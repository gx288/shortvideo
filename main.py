import os
import re
import requests
import subprocess
import textwrap
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip
import numpy as np
import glob
from google.cloud import texttospeech
import gspread
from google.oauth2.service_account import Credentials
import random
import unicodedata

# Configuration options (modify these as needed)
NUM_VIDEOS_TO_CREATE = 3  # Number of videos to create per run
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
        print("  ffmpeg is available")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("  Warning: ffmpeg not found. Some audio processing may fail.")
        return False

# Hàm xử lý tên file để loại bỏ dấu và ký tự đặc biệt
def clean_filename(text, max_length=50):
    # Chuẩn hóa Unicode: chuyển các ký tự có dấu thành không dấu
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    # Thay khoảng trắng bằng dấu gạch dưới
    text = text.replace(' ', '_')
    # Chỉ giữ chữ cái, số, dấu gạch dưới, và dấu gạch ngang
    text = re.sub(r'[^\w-]', '', text)
    # Loại bỏ các dấu gạch dưới liên tiếp
    text = re.sub(r'_+', '_', text)
    # Cắt ngắn tên file nếu quá dài
    text = text[:max_length].strip('_')
    # Nếu tên rỗng hoặc chỉ chứa ký tự không hợp lệ, trả về tên mặc định
    if not text or text == '':
        text = f"video_{random.randint(1000, 9999)}"
    return text.lower()

# Directory setup
output_dir = "output"
print("Stage 1: Creating output directory...")
os.makedirs(output_dir, exist_ok=True)
print(f"  Created directory: {output_dir}")

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
        print(f"  Error: Worksheet '{worksheet_name}' not found. Skipping.")
        continue

    # Read rows from the worksheet
    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break
        if i == 0:  # Skip header
            continue
        if len(row) > 7 and (not row[7] or row[7].strip() == ''):
            print(f"  Processing row {i + 1} in worksheet '{worksheet_name}'...")
            selected_row = row
            selected_row_num = i + 1  # 1-based index

            # Extract and clean title for filename
            raw_content = selected_row[1] if len(selected_row) > 1 else ''
            raw_content = re.sub(r'\*+', '', raw_content)  # Remove asterisks
            raw_content = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_content)  # Remove emojis
            raw_content = re.sub(r'#\w+\s*', '', raw_content)  # Remove hashtags
            lines = [line.strip() for line in raw_content.split('\n') if line.strip()]
            title_text = lines[0].replace('Tiêu đề:', '').strip() if lines else 'Untitled'
            content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text

            clean_title = clean_filename(title_text)
            print(f"  Original title: {title_text}")
            print(f"  Clean title for filename: {clean_title}")
            print(f"  Clean content length: {len(content_text)} chars")

            # Save clean_title for update_sheet.py
            with open(os.path.join(output_dir, "clean_title.txt"), "w") as f:
                f.write(clean_title)

            # Extract cover image URL from column D (index 3)
            bg_image_url = selected_row[3] if len(selected_row) > 3 else 'https://via.placeholder.com/1080x1920?text=No+Image'
            print(f"  Cover image URL: {bg_image_url}")

            # Stage 2: Create audio with Google Cloud TTS
            print("Stage 2: Creating audio with Google Cloud TTS...")
            try:
                client = texttospeech.TextToSpeechClient.from_service_account_file('google_tts_key.json')
                synthesis_input = texttospeech.SynthesisInput(text=content_text)
                voice = texttospeech.VoiceSelectionParams(
                    language_code="vi-VN",
                    name="vi-VN-Wavenet-C"  # Changed from vi-VN-Wavenet-A to vi-VN-Wavenet-C
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
                print(f"  Saved audio at: {audio_path}")
            except Exception as e:
                print(f"  Error creating audio: {e}. Skipping row {selected_row_num}.")
                continue

            # Cut audio to max 55s using ffmpeg
            try:
                temp_audio = os.path.join(output_dir, "temp_voiceover.mp3")
                subprocess.run([
                    "ffmpeg", "-i", audio_path, "-t", "55", "-c:a", "mp3", "-b:a", "96k", temp_audio
                ], check=True, capture_output=True)
                os.replace(temp_audio, audio_path)
                print("  Cut audio to 55s with ffmpeg")
            except Exception as e:
                print(f"  Warning: Failed to cut audio with ffmpeg: {e}. Using original audio.")

            # Stage 3: Create title image
            def create_title_image(title, bg_image_url, output_path):
                print("Stage 3: Creating title image...")
                try:
                    bg_image = Image.open(requests.get(bg_image_url, stream=True, timeout=10).raw).convert("RGB")
                    print("  Downloaded background image.")
                except Exception as e:
                    print(f"  Warning: Failed to download background image: {e}. Using black background.")
                    bg_image = Image.new("RGB", (720, 1280), (0, 0, 0))

                target_size = (720, 1280)
                img_ratio = bg_image.width / bg_image.height
                target_ratio = target_size[0] / target_size[1]

                if img_ratio > target_ratio:
                    new_height = target_size[1]
                    new_width = int(new_height * img_ratio)
                else:
                    new_width = target_size[0]
                    new_height = int(new_width / img_ratio)

                bg_image = bg_image.resize((new_width, new_height), Image.LANCZOS)
                final_image = Image.new("RGB", target_size, (0, 0, 0))
                paste_x = (target_size[0] - new_width) // 2
                paste_y = (target_size[1] - new_height) // 2
                final_image.paste(bg_image, (paste_x, paste_y))

                draw = ImageDraw.Draw(final_image)
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
                        print(f"  Using font: {font_path}")
                        break
                    except:
                        continue
                if not font:
                    print("  Error: No custom font found. Default font is not suitable for font_size = 80. Please install a font like Arial or Roboto.")
                    return

                max_width = 576
                min_height = 400
                max_height = 768
                target_height = 533
                line_spacing = 15
                wrap_width = 30

                def get_text_dimensions(text, font, wrap_width):
                    wrapped_text = []
                    current_line = ""
                    for word in text.split():
                        test_line = current_line + (" " if current_line else "") + word
                        text_bbox = draw.textbbox((0, 0), test_line, font=font)
                        text_width = text_bbox[2] - text_bbox[0]
                        if text_width <= max_width:
                            current_line = test_line
                        else:
                            if current_line:
                                wrapped_text.append(current_line)
                            current_line = word
                    if current_line:
                        wrapped_text.append(current_line)

                    total_height = 0
                    max_text_width = 0
                    for line in wrapped_text:
                        text_bbox = draw.textbbox((0, 0), line, font=font)
                        text_width = text_bbox[2] - text_bbox[0]
                        text_height = text_bbox[3] - text_bbox[1]
                        max_text_width = max(max_text_width, text_width)
                        total_height += text_height + line_spacing
                    return wrapped_text, max_text_width, total_height - line_spacing

                wrapped_text, max_text_width, total_height = get_text_dimensions(title, font, wrap_width)

                max_attempts = 20
                attempt = 0
                while (total_height < min_height or total_height > max_height) and wrap_width >= 10 and attempt < max_attempts:
                    if total_height < min_height:
                        wrap_width -= 1
                    elif total_height > max_height:
                        wrap_width += 1

                    if wrap_width < 10:
                        print("  Warning: wrap_width reached minimum (10). Cannot increase lines further.")
                        break

                    wrapped_text, max_text_width, total_height = get_text_dimensions(title, font, wrap_width)
                    print(f"  Adjusted wrap_width: {wrap_width}, max_text_width: {max_text_width}, total_height: {total_height}, lines: {len(wrapped_text)}")
                    attempt += 1

                print(f"  Final font_size: {font_size}, wrap_width: {wrap_width}, max_text_width: {max_text_width}, total_height: {total_height}, lines: {len(wrapped_text)}")

                text_area_height = total_height + 40
                text_area = Image.new("RGBA", (720, text_area_height), (0, 0, 0, int(255 * 0.7)))
                text_draw = ImageDraw.Draw(text_area)

                current_y = 20
                for line in wrapped_text:
                    text_bbox = text_draw.textbbox((0, 0), line, font=font)
                    text_width = text_bbox[2] - text_bbox[0]
                    text_x = (720 - text_width) // 2
                    text_draw.text((text_x, current_y), line, font=font, fill=(255, 255, 255), stroke_width=2, stroke_fill=(0, 0, 0))
                    current_y += (text_bbox[3] - text_bbox[1]) + line_spacing

                text_y = (1280 - text_area_height) // 2
                final_image = final_image.convert("RGBA")
                final_image.paste(text_area, (0, text_y), text_area)
                final_image = final_image.convert("RGB")

                final_image.save(output_path)
                print(f"  Saved title image at: {output_path}")

            title_image_path = os.path.join(output_dir, "title_image.jpg")
            create_title_image(title_text, bg_image_url, title_image_path)
            if not os.path.exists(title_image_path):
                print(f"  Error: Failed to create title image for row {selected_row_num}. Skipping.")
                continue

            # Stage 4: Download images
            def download_images_with_icrawler(keyword, num_images, output_dir):
                print("Stage 4: Attempting to download images...")
                keyword_clean = clean_filename(keyword)  # Clean keyword for directory
                keyword_dir = os.path.join(output_dir, keyword_clean)
                os.makedirs(keyword_dir, exist_ok=True)

                try:
                    from icrawler.builtin import GoogleImageCrawler
                    google_crawler = GoogleImageCrawler(storage={'root_dir': keyword_dir})
                    google_crawler.crawl(keyword=keyword, max_num=num_images, min_size=(500, 500))
                    print(f"  Crawled images for keyword: {keyword}")
                except ImportError:
                    print("  Warning: icrawler not installed. Using fallback images.")
                    return [title_image_path] * max(2, num_images)
                except Exception as e:
                    print(f"  Warning: Image crawling failed: {e}. Using fallback images.")
                    return [title_image_path] * max(2, num_images)

                image_pattern = os.path.join(keyword_dir, "*.jpg")
                downloaded_files = glob.glob(image_pattern)
                image_paths = []

                if len(downloaded_files) < 2:
                    print("  Warning: Not enough images downloaded. Using fallback.")
                    return [title_image_path] * max(2, num_images)

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
                        print(f"  Warning: Failed to process image {full_path}: {e}")
                        continue

                if not image_paths:
                    print("  Warning: No valid images processed. Using fallback.")
                    return [title_image_path] * max(2, num_images)

                return image_paths

            keyword = title_text[:50]
            additional_images = download_images_with_icrawler(keyword, 10, output_dir)
            image_paths = [title_image_path] + additional_images
            print(f"  Retrieved {len(additional_images)} images")

            # Stage 5: Create video with varied transitions
            def create_video(image_paths, audio_path, output_path):
                print("Stage 5: Creating video...")
                try:
                    audio = AudioFileClip(audio_path)
                    audio_duration = audio.duration
                except Exception as e:
                    print(f"  Error loading audio {audio_path}: {e}. Skipping row {selected_row_num}.")
                    return False

                clips = []
                num_images = len(image_paths)
                title_duration = audio_duration * 1.2 / num_images
                other_duration = (audio_duration - title_duration) / (num_images - 1) if num_images > 1 else audio_duration

                def zoom_in(t, duration):
                    return 1 + 0.2 * (t / duration)

                def zoom_out(t, duration):
                    return 1.2 - 0.2 * (t / duration)

                def pan_left(t, duration):
                    return (0.2 * (t / duration) * 720, 'center')

                def pan_right(t, duration):
                    return (-0.2 * (t / duration) * 720, 'center')

                def pan_up(t, duration):
                    return ('center', 0.2 * (t / duration) * 1280)

                def pan_down(t, duration):
                    return ('center', -0.2 * (t / duration) * 1280)

                transitions = [zoom_in, zoom_out, pan_left, pan_right, pan_up, pan_down]

                for i, img_path in enumerate(image_paths):
                    try:
                        duration = title_duration if i == 0 else other_duration
                        clip = ImageClip(img_path, duration=duration)
                        transition = transitions[i % len(transitions)]
                        clip = clip.resize(lambda t: transition(t, duration) if transition in [zoom_in, zoom_out] else 1.0).set_position(lambda t: transition(t, duration) if transition not in [zoom_in, zoom_out] else 'center')
                        clips.append(clip)
                        print(f"  Applied transition {transition.__name__} to {img_path}")
                    except Exception as e:
                        print(f"  Warning: Failed to process image {img_path}: {e}")
                        continue

                if not clips:
                    print(f"  Error: No valid clips to create video for row {selected_row_num}. Skipping.")
                    return False

                try:
                    video = concatenate_videoclips(clips, method="compose")
                    video = video.set_audio(audio)
                    video.write_videofile(output_path, codec="libx265", audio_codec="aac", fps=15, bitrate="1000k", audio_bitrate="96k", ffmpeg_params=["-preset", "medium"])
                    print(f"  Saved video at: {output_path}")
                    return True
                except Exception as e:
                    print(f"  Error saving video: {e}. Skipping row {selected_row_num}.")
                    return False

            output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
            if create_video(image_paths, audio_path, output_video_path):
                videos_created += 1
                print(f"Video created successfully at: {output_video_path}")
                print(f"Video size: {os.path.getsize(output_video_path) / (1024 * 1024):.2f} MB")
            else:
                print(f"Failed to create video for row {selected_row_num} in worksheet '{worksheet_name}'.")

            # Clean up temporary files (except video)
            print("Cleaning up temporary files...")
            keyword_clean = clean_filename(keyword)
            for file in glob.glob(os.path.join(output_dir, keyword_clean, "*.jpg")):
                try:
                    os.remove(file)
                except:
                    pass
            if os.path.exists(audio_path):
                try:
                    os.remove(audio_path)
                except:
                    pass
            if os.path.exists(title_image_path):
                try:
                    os.remove(title_image_path)
                except:
                    pass
            print("Cleanup complete.")

if videos_created == 0:
    print("No videos were created. Exiting.")
    exit(1)

print(f"Successfully created {videos_created} video(s).")
