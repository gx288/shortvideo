import os
import re
import requests
import subprocess
import textwrap
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip, CompositeVideoClip, TextClip
import numpy as np
import glob
from google.cloud import texttospeech
import gspread
from google.oauth2.service_account import Credentials
import random
import unicodedata
import sys

# Configuration options
NUM_VIDEOS_TO_CREATE = 1
WORKSHEET_LIST = [
    "Phòng mạch",
    "Sheet2",
    "Sheet3",
]

# Fallback for ANTIALIAS
Image.ANTIALIAS = Image.LANCZOS

# Check for ffmpeg
def check_ffmpeg():
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True)
        print("  ffmpeg is available")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("  ERROR: ffmpeg not found. Install ffmpeg to continue.")
        return False

# Clean filename
def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    if not text:
        text = f"video_{random.randint(1000, 9999)}"
    return text.lower()

# Directory setup
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)
print(f"Stage 0: Output directory ready: {output_dir}")

# Google Sheets setup
print("Stage 0: Initializing Google Sheets...")
try:
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
    gc = gspread.authorize(creds)
    SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'
except Exception as e:
    print(f"  ERROR: Failed to authenticate Google Sheets: {e}")
    sys.exit(1)

videos_created = 0

# Main loop
for worksheet_name in WORKSHEET_LIST:
    if videos_created >= NUM_VIDEOS_TO_CREATE:
        break

    print(f"\nChecking worksheet: {worksheet_name}")
    try:
        worksheet = gc.open_by_key(SHEET_ID).worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  ERROR: Worksheet '{worksheet_name}' not found. Skipping.")
        continue
    except Exception as e:
        print(f"  ERROR: Failed to open worksheet: {e}")
        sys.exit(1)

    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break
        if i == 0:  # Skip header
            continue
        if len(row) > 7 and (not row[7] or row[7].strip() == ''):
            selected_row = row
            selected_row_num = i + 1
            print(f"\nProcessing row {selected_row_num} in '{worksheet_name}'...")

            try:
                # === EXTRACT DATA ===
                raw_content = selected_row[1] if len(selected_row) > 1 else ''
                raw_content = re.sub(r'\*+', '', raw_content)
                raw_content = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_content)
                raw_content = re.sub(r'#\w+\s*', '', raw_content)
                lines = [line.strip() for line in raw_content.split('\n') if line.strip()]
                title_text = lines[0].replace('Tiêu đề:', '').strip() if lines else 'Untitled'
                content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text

                clean_title = clean_filename(title_text)
                print(f"  Title: {title_text}")
                print(f"  Clean filename: {clean_title}")

                # Save clean_title
                with open(os.path.join(output_dir, "clean_title.txt"), "w", encoding="utf-8") as f:
                    f.write(clean_title)

                bg_image_url = selected_row[3] if len(selected_row) > 3 else 'https://via.placeholder.com/1080x1920?text=No+Image'
                print(f"  Cover image URL: {bg_image_url}")

                # === STAGE 1: DOWNLOAD IMAGES FIRST ===
                print("Stage 1: Downloading images...")
                keyword = title_text[:50]
                image_paths = []

                # Download cover image
                try:
                    cover_path = os.path.join(output_dir, "cover.jpg")
                    response = requests.get(bg_image_url, timeout=10)
                    response.raise_for_status()
                    with open(cover_path, "wb") as f:
                        f.write(response.content)
                    # Resize to 720x1280
                    img = Image.open(cover_path).convert("RGB")
                    img = resize_and_crop(img, (720, 1280))
                    img.save(cover_path)
                    image_paths.append(cover_path)
                    print(f"  Downloaded cover: {cover_path}")
                except Exception as e:
                    print(f"  ERROR: Failed to download cover image: {e}")
                    sys.exit(1)

                # Download additional images
                try:
                    from icrawler.builtin import GoogleImageCrawler
                    keyword_clean = clean_filename(keyword)
                    keyword_dir = os.path.join(output_dir, keyword_clean)
                    os.makedirs(keyword_dir, exist_ok=True)

                    google_crawler = GoogleImageCrawler(storage={'root_dir': keyword_dir})
                    google_crawler.crawl(keyword=keyword, max_num=10, min_size=(500, 500))
                    print(f"  Crawled images for: {keyword}")

                    downloaded = glob.glob(os.path.join(keyword_dir, "*.jpg"))
                    for img_file in downloaded[:9]:
                        try:
                            img = Image.open(img_file).convert("RGB")
                            img = resize_and_crop(img, (720, 1280))
                            img.save(img_file)
                            image_paths.append(img_file)
                        except Exception as e:
                            print(f"  Warning: Failed to process {img_file}: {e}")
                            continue
                except ImportError:
                    print("  WARNING: icrawler not installed. Using cover only.")
                except Exception as e:
                    print(f"  WARNING: Crawling failed: {e}. Using cover only.")

                if len(image_paths) < 2:
                    print("  WARNING: Less than 2 images. Duplicating cover.")
                    image_paths = image_paths + [image_paths[0]] * (10 - len(image_paths))

                print(f"  Total images ready: {len(image_paths)}")

                # === STAGE 2: CREATE AUDIO (AFTER IMAGES) ===
                print("Stage 2: Creating audio with Google Cloud TTS...")
                audio_path = os.path.join(output_dir, "voiceover.mp3")
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
                    response = client.synthesize_speech(input=synthesis_input, voice=voice, audio_config=audio_config)
                    with open(audio_path, "wb") as f:
                        f.write(response.audio_content)
                    print(f"  Audio saved: {audio_path}")

                    # Cut to 55s
                    temp_audio = audio_path + ".temp.mp3"
                    subprocess.run([
                        "ffmpeg", "-i", audio_path, "-t", "55", "-c:a", "mp3", "-b:a", "96k", temp_audio
                    ], check=True, capture_output=True)
                    os.replace(temp_audio, audio_path)
                    print("  Audio cut to 55s")
                except Exception as e:
                    print(f"  ERROR: TTS failed: {e}")
                    sys.exit(1)

                # === STAGE 3: CREATE VIDEO WITH TITLE OVERLAY FULL VIDEO ===
                print("Stage 3: Creating video with title overlay...")
                try:
                    audio = AudioFileClip(audio_path)
                    total_duration = min(audio.duration, 55)

                    clips = []
                    num_images = len(image_paths)
                    duration_per_clip = total_duration / num_images

                    transitions = [
                        lambda t, d: 1 + 0.2 * (t / d),  # zoom in
                        lambda t, d: 1.2 - 0.2 * (t / d),  # zoom out
                        lambda t, d: (0.2 * (t / d) * 720, 'center'),  # pan left
                        lambda t, d: (-0.2 * (t / d) * 720, 'center'),  # pan right
                        lambda t, d: ('center', 0.2 * (t / d) * 1280),  # pan up
                        lambda t, d: ('center', -0.2 * (t / d) * 1280),  # pan down
                    ]

                    for idx, img_path in enumerate(image_paths):
                        clip = ImageClip(img_path).set_duration(duration_per_clip)
                        trans = transitions[idx % len(transitions)]
                        if trans.__name__ in ["zoom_in", "zoom_out"]:
                            clip = clip.resize(trans)
                        else:
                            clip = clip.set_position(trans)
                        clips.append(clip)

                    video = concatenate_videoclips(clips, method="compose")

                    # === ADD TITLE TEXT OVERLAY FULL VIDEO ===
                    txt_clip = TextClip(
                        title_text,
                        fontsize=70,
                        color='white',
                        font='Arial-Bold',
                        stroke_color='black',
                        stroke_width=2,
                        size=(576, None),
                        method='caption',
                        align='center'
                    ).set_position(('center', 'center')).set_duration(total_duration)

                    # Background semi-transparent
                    bg_txt = TextClip(
                        "", color='black', size=(620, txt_clip.h + 40)
                    ).set_opacity(0.6).set_position(('center', 'center')).set_duration(total_duration)

                    final_video = CompositeVideoClip([video, bg_txt, txt_clip])
                    final_video = final_video.set_audio(audio)

                    output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
                    final_video.write_videofile(
                        output_video_path,
                        codec="libx265",
                        audio_codec="aac",
                        fps=15,
                        bitrate="700k",
                        audio_bitrate="96k",
                        ffmpeg_params=["-preset", "medium"],
                        threads=4
                    )
                    print(f"  VIDEO CREATED: {output_video_path}")
                    print(f"  Size: {os.path.getsize(output_video_path) / (1024*1024):.2f} MB")

                    videos_created += 1

                    # === UPDATE SHEET COLUMN H (index 7) ===
                    try:
                        worksheet.update_cell(selected_row_num, 8, "DONE")
                        print(f"  Updated row {selected_row_num} as DONE")
                    except Exception as e:
                        print(f"  WARNING: Failed to update sheet: {e}")

                except Exception as e:
                    print(f"  ERROR: Video creation failed: {e}")
                    sys.exit(1)

                # === CLEANUP ===
                print("Cleaning up...")
                for f in [audio_path] + glob.glob(os.path.join(output_dir, "cover.jpg")):
                    if os.path.exists(f):
                        os.remove(f)
                for d in glob.glob(os.path.join(output_dir, clean_filename(keyword))):
                    if os.path.isdir(d):
                        for file in glob.glob(os.path.join(d, "*")):
                            os.remove(file)
                        os.rmdir(d)
                print("Cleanup done.\n")

            except Exception as e:
                print(f"  FATAL ERROR at row {selected_row_num}: {e}")
                sys.exit(1)

# Final
if videos_created == 0:
    print("No videos created.")
    sys.exit(1)
else:
    print(f"SUCCESS: {videos_created} video(s) created.")
    sys.exit(0)
