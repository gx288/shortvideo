import os
import re
import requests
import subprocess
import textwrap
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip, CompositeVideoClip
import numpy as np
import glob
from google.cloud import texttospeech
import gspread
from google.oauth2.service_account import Credentials
import random
import unicodedata

# Configuration options (modify these as needed)
NUM_VIDEOS_TO_CREATE = 1  # Number of videos to create per run
WORKSHEET_LIST = [
    "Phòng mạch",
    "Sheet2",
    "Sheet3",
    # Add more sheet names here
]

# Fallback for ANTIALIAS in Pillow
Image.ANTIALIAS = Image.LANCZOS

# Check for ffmpeg
def check_ffmpeg():
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True)
        print("ffmpeg is available")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("Warning: ffmpeg not found. Some audio processing may fail.")
        return False

# Hàm xử lý tên file để loại bỏ dấu và ký tự đặc biệt
def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    if not text or text == '':
        text = f"video_{random.randint(1000, 9999)}"
    return text.lower()

# Directory setup
output_dir = "output"
print("Stage 1: Creating output directory...")
os.makedirs(output_dir, exist_ok=True)
print(f"Created directory: {output_dir}")

# Google Sheets setup
print("Stage 0: Initializing Google Sheets...")
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'

# Counter for created videos
videos_created = 0

# Process up to NUM_VIDEOS_TO_CREATE videos
for worksheet_name in WORKSHEET_LIST:
    if videos_created >= NUM_VIDEOS_TO_CREATE:
        break
    print(f"\nChecking worksheet: {worksheet_name}")
    try:
        worksheet = gc.open_by_key(SHEET_ID).worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Error: Worksheet '{worksheet_name}' not found. Skipping.")
        continue

    # Read rows from the worksheet
    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break
        if i == 0:  # Skip header
            continue
        if len(row) > 7 and (not row[7] or row[7].strip() == ''):
            print(f"Processing row {i + 1} in worksheet '{worksheet_name}'...")
            selected_row = row
            selected_row_num = i + 1  # 1-based index

            # Extract and clean title for filename
            raw_content = selected_row[1] if len(selected_row) > 1 else ''
            raw_content = re.sub(r'\*+', '', raw_content)
            raw_content = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_content)
            raw_content = re.sub(r'#\w+\s*', '', raw_content)
            lines = [line.strip() for line in raw_content.split('\n') if line.strip()]
            title_text = lines[0].replace('Tiêu đề:', '').strip() if lines else 'Untitled'
            content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text
            clean_title = clean_filename(title_text)
            print(f"Original title: {title_text}")
            print(f"Clean title for filename: {clean_title}")
            print(f"Clean content length: {len(content_text)} chars")

            # Save clean_title for update_sheet.py
            with open(os.path.join(output_dir, "clean_title.txt"), "w", encoding='utf-8') as f:
                f.write(clean_title)

            # Extract cover image URL from column D (index 3)
            bg_image_url = selected_row[3] if len(selected_row) > 3 else 'https://via.placeholder.com/1080x1920?text=No+Image'
            print(f"Cover image URL: {bg_image_url}")

            # Stage 2: Create audio with Google Cloud TTS
            print("Stage 2: Creating audio with Google Cloud TTS...")
            try:
                client = texttospeech.TextToSpeechClient.from_service_account_file('google_tts_key.json')
                synthesis_input = texttospeech.SynthesisInput(text=content_text)
                voice = texttospeech.VoiceSelectionParams(
                    language_code="vi-VN",
                    name="vi-VN-Wavenet-C"
                )
                audio_config = texttospeech.AudioConfig(
                    audio_encoding=texttospeech.AudioEncoding.MP3,
                    speaking_rate=1.25,
                    pitch=0.0,
                    sample_rate_hertz=44100
                )
                response = client.synthesize_speech(
                    input=synthesis_input, voice=voice, audio_config=audio_config
                )
                audio_path = os.path.join(output_dir, "voiceover.mp3")
                with open(audio_path, "wb") as out:
                    out.write(response.audio_content)
                print(f"Saved audio at: {audio_path}")
            except Exception as e:
                print(f"Error creating audio: {e}. Skipping row {selected_row_num}.")
                continue

            # Cut audio to max 55s using ffmpeg
            try:
                temp_audio = os.path.join(output_dir, "temp_voiceover.mp3")
                subprocess.run([
                    "ffmpeg", "-i", audio_path, "-t", "55", "-c:a", "mp3", "-b:a", "96k", temp_audio
                ], check=True, capture_output=True)
                os.replace(temp_audio, audio_path)
                print("Cut audio to 55s with ffmpeg")
            except Exception as e:
                print(f"Warning: Failed to cut audio with ffmpeg: {e}. Using original audio.")

            # Stage 4: Download images (skip title image)
            def download_images_with_icrawler(keyword, num_images, output_dir):
                print("Stage 4: Attempting to download images...")
                keyword_clean = clean_filename(keyword)
                keyword_dir = os.path.join(output_dir, keyword_clean)
                os.makedirs(keyword_dir, exist_ok=True)
                try:
                    from icrawler.builtin import GoogleImageCrawler
                    google_crawler = GoogleImageCrawler(storage={'root_dir': keyword_dir})
                    google_crawler.crawl(keyword=keyword, max_num=num_images, min_size=(500, 500))
                    print(f"Crawled images for keyword: {keyword}")
                except ImportError:
                    print("Warning: icrawler not installed. Using fallback.")
                    return []
                except Exception as e:
                    print(f"Warning: Image crawling failed: {e}. Using fallback.")
                    return []

                image_pattern = os.path.join(keyword_dir, "*.jpg")
                downloaded_files = glob.glob(image_pattern)
                image_paths = []
                for full_path in downloaded_files[:num_images]:
                    try:
                        img = Image.open(full_path).convert("RGB")
                        img_ratio = img.width / img.height
                        target_ratio = 720 / 1280
                        if img_ratio > target_ratio:
                            new_height = 1280
                            new_width = int(new_height * img_ratio)
                        else:
                            new_width = 720
                            new_height = int(new_width / img_ratio)
                        img = img.resize((new_width, new_height), Image.LANCZOS)
                        final_image = Image.new("RGB", (720, 1280), (0, 0, 0))
                        paste_x = (720 - new_width) // 2
                        paste_y = (1280 - new_height) // 2
                        final_image.paste(img, (paste_x, paste_y))
                        final_image.save(full_path)
                        image_paths.append(full_path)
                    except Exception as e:
                        print(f"Warning: Failed to process image {full_path}: {e}")
                        continue
                return image_paths

            keyword = title_text[:50]
            additional_images = download_images_with_icrawler(keyword, 10, output_dir)
            if not additional_images:
                print("Warning: No images downloaded. Creating fallback image.")
                fallback_path = os.path.join(output_dir, "fallback.jpg")
                Image.new("RGB", (720, 1280), (30, 30, 50)).save(fallback_path)
                additional_images = [fallback_path] * 3

            print(f"Retrieved {len(additional_images)} images")

            # Stage 5: Create video with title overlay throughout
            def create_video(image_paths, audio_path, output_path, title_text):
                print("Stage 5: Creating video with persistent title overlay...")
                try:
                    audio = AudioFileClip(audio_path)
                    audio_duration = min(audio.duration, 55)
                except Exception as e:
                    print(f"Error loading audio {audio_path}: {e}. Skipping row.")
                    return False

                clips = []
                num_images = len(image_paths)
                duration_per_image = audio_duration / num_images if num_images > 0 else 5.0

                # Load font
                font_size = 80
                font_paths = [
                    "Roboto-Bold.ttf",
                    "C:/Windows/Fonts/Arialbd.ttf",
                    "C:/Windows/Fonts/Calibrib.ttf",
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
                ]
                font = None
                for font_path in font_paths:
                    try:
                        font = ImageFont.truetype(font_path, font_size)
                        print(f"Using font: {font_path}")
                        break
                    except:
                        continue
                if not font:
                    print("Warning: No font found. Using default.")
                    font = ImageFont.load_default()

                # Wrap text
                max_width = 576
                line_spacing = 15

                def wrap_text(text, font, max_width):
                    words = text.split()
                    lines = []
                    current_line = ""
                    for word in words:
                        test_line = current_line + (" " if current_line else "") + word
                        bbox = font.getbbox(test_line)
                        text_width = bbox[2] - bbox[0]
                        if text_width <= max_width:
                            current_line = test_line
                        else:
                            if current_line:
                                lines.append(current_line)
                            current_line = word
                    if current_line:
                        lines.append(current_line)
                    return lines

                wrapped_lines = wrap_text(title_text, font, max_width)

                # Create title overlay
                def create_title_overlay():
                    line_heights = [font.getbbox(line)[3] - font.getbbox(line)[1] for line in wrapped_lines]
                    total_text_height = sum(line_heights) + line_spacing * (len(wrapped_lines) - 1)
                    padding = 40
                    overlay_height = total_text_height + padding
                    overlay = Image.new("RGBA", (720, 1280), (0, 0, 0, 0))
                    draw = ImageDraw.Draw(overlay)
                    current_y = (1280 - overlay_height) // 2 + 20
                    for line in wrapped_lines:
                        bbox = draw.textbbox((0, 0), line, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_x = (720 - text_width) // 2
                        bg_padding = 20
                        draw.rectangle([
                            (text_x - bg_padding, current_y - 10),
                            (text_x + text_width + bg_padding, current_y + (bbox[3] - bbox[1]) + 10)
                        ], fill=(0, 0, 0, 180))
                        draw.text((text_x, current_y), line, font=font, fill=(255, 255, 255),
                                  stroke_width=2, stroke_fill=(0, 0, 0))
                        current_y += (bbox[3] - bbox[1]) + line_spacing
                    return overlay

                title_overlay_img = create_title_overlay()
                title_overlay_clip = ImageClip(np.array(title_overlay_img)).set_duration(audio_duration).set_position('center')

                # Transitions
                def zoom_in(t, duration): return 1 + 0.2 * (t / duration)
                def zoom_out(t, duration): return 1.2 - 0.2 * (t / duration)
                def pan_left(t, duration): return (0.2 * (t / duration) * 720, 'center')
                def pan_right(t, duration): return (-0.2 * (t / duration) * 720, 'center')
                def pan_up(t, duration): return ('center', 0.2 * (t / duration) * 1280)
                def pan_down(t, duration): return ('center', -0.2 * (t / duration) * 1280)
                transitions = [zoom_in, zoom_out, pan_left, pan_right, pan_up, pan_down]

                for i, img_path in enumerate(image_paths):
                    try:
                        img = Image.open(img_path).convert("RGB")
                        img_ratio = img.width / img.height
                        target_ratio = 720 / 1280
                        if img_ratio > target_ratio:
                            new_height = 1280
                            new_width = int(new_height * img_ratio)
                        else:
                            new_width = 720
                            new_height = int(new_width / img_ratio)
                        img = img.resize((new_width, new_height), Image.LANCZOS)
                        final_img = Image.new("RGB", (720, 1280), (0, 0, 0))
                        paste_x = (720 - new_width) // 2
                        paste_y = (1280 - new_height) // 2
                        final_img.paste(img, (paste_x, paste_y))

                        clip = ImageClip(np.array(final_img), duration=duration_per_image)
                        transition = transitions[i % len(transitions)]
                        if transition in [zoom_in, zoom_out]:
                            clip = clip.resize(lambda t: transition(t, duration_per_image)).set_position('center')
                        else:
                            clip = clip.set_position(lambda t: transition(t, duration_per_image))

                        start_time = i * duration_per_image
                        end_time = (i + 1) * duration_per_image
                        overlay_part = title_overlay_clip.subclip(start_time, end_time).set_position('center')

                        composite = CompositeVideoClip([clip.set_duration(duration_per_image), overlay_part])
                        clips.append(composite)
                        print(f"Applied {transition.__name__} to image {i+1}/{num_images}")
                    except Exception as e:
                        print(f"Warning: Failed to process image {img_path}: {e}")
                        continue

                if not clips:
                    print("Error: No valid clips to create video.")
                    return False

                try:
                    final_video = concatenate_videoclips(clips, method="compose")
                    final_video = final_video.set_audio(audio)
                    final_video.write_videofile(
                        output_path,
                        codec="libx265",
                        audio_codec="aac",
                        fps=15,
                        bitrate="700k",
                        audio_bitrate="96k",
                        ffmpeg_params=["-preset", "medium"],
                        threads=4,
                        logger=None
                    )
                    print(f"Saved video with persistent title at: {output_path}")
                    return True
                except Exception as e:
                    print(f"Error saving video: {e}")
                    return False

            output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
            if create_video(additional_images, audio_path, output_video_path, title_text):
                videos_created += 1
                print(f"Video created successfully at: {output_video_path}")
                print(f"Video size: {os.path.getsize(output_video_path) / (1024 * 1024):.2f} MB")
            else:
                print(f"Failed to create video for row {selected_row_num} in worksheet '{worksheet_name}'.")

            # Clean up temporary files
            print("Cleaning up temporary files...")
            keyword_clean = clean_filename(keyword)
            for file in glob.glob(os.path.join(output_dir, keyword_clean, "*.*")):
                try:
                    os.remove(file)
                except:
                    pass
            if os.path.exists(audio_path):
                try:
                    os.remove(audio_path)
                except:
                    pass
            print("Cleanup complete.")

if videos_created == 0:
    print("No videos were created. Exiting.")
    exit(1)

print(f"Successfully created {videos_created} video(s).")
