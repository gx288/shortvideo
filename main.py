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

# Directory setup
output_dir = "output"
print("Stage 1: Creating output directory...")
os.makedirs(output_dir, exist_ok=True)
print(f"  Created directory: {output_dir}")

# Stage 0: Read from Google Sheets
print("Stage 0: Reading from Google Sheets...")
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)

SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'
worksheet = gc.open_by_key(SHEET_ID).worksheet('Phòng mạch')

rows = worksheet.get_all_values()
selected_row = None
selected_row_num = None
for i, row in enumerate(rows):
    if i == 0:  # Skip header
        continue
    if len(row) > 7 and (not row[7] or row[7].strip() == ''):
        selected_row = row
        selected_row_num = i + 1  # 1-based index
        break

if not selected_row:
    print("  No unprocessed rows found. Skipping video creation.")
    exit(0)

print(f"  Processing row {selected_row_num}...")
# Extract and clean title for filename
raw_content = selected_row[1] if len(selected_row) > 1 else ''
raw_content = re.sub(r'\*+', '', raw_content)  # Remove asterisks
raw_content = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_content)  # Remove emojis
raw_content = re.sub(r'#\w+\s*', '', raw_content)  # Remove hashtags
lines = [line.strip() for line in raw_content.split('\n') if line.strip()]
title_text = lines[0].replace('Tiêu đề:', '').strip() if lines else 'Untitled'
content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text

# Clean title for filename
clean_title = re.sub(r'[^\w\s-]', '', title_text.replace(' ', '_'))[:50]  # Remove special chars, limit length
clean_title = clean_title.lower() if clean_title else f"video_{selected_row_num}"
print(f"  Clean title: {title_text}")
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
        name="vi-VN-Wavenet-A"
    )
    audio_config = texttospeech.AudioConfig(
        audio_encoding=texttospeech.AudioEncoding.MP3,
        speaking_rate=1.25,
        pitch=0.0
    )
    response = client.synthesize_speech(
        input=synthesis_input, voice=voice, audio_config=audio_config
    )
    audio_path = os.path.join(output_dir, "voiceover.mp3")
    with open(audio_path, "wb") as out:
        out.write(response.audio_content)
    print(f"  Saved audio at: {audio_path}")
except Exception as e:
    print(f"  Error creating audio: {e}. Exiting.")
    exit(1)

# Stage 3: Create title image
def create_title_image(title, bg_image_url, output_path):
    print("Stage 3: Creating title image...")
    try:
        bg_image = Image.open(requests.get(bg_image_url, stream=True, timeout=10).raw).convert("RGB")
        print("  Downloaded background image.")
    except Exception as e:
        print(f"  Warning: Failed to download background image: {e}. Using black background.")
        bg_image = Image.new("RGB", (1080, 1920), (0, 0, 0))

    target_size = (1080, 1920)
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
    font_size = 120  # Reduced from 150 to ~80%
    font_paths = [
        "Roboto-Bold.ttf",
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
        print("  Warning: No custom font found. Using default font.")
        font = ImageFont.load_default()

    max_width = 918
    max_height = 960
    min_height = 768
    line_spacing = 20

    def get_text_dimensions(text, font, wrap_width):
        wrapped_text = textwrap.wrap(text, width=wrap_width)
        total_height = 0
        max_text_width = 0
        for line in wrapped_text:
            text_bbox = draw.textbbox((0, 0), line, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]
            max_text_width = max(max_text_width, text_width)
            total_height += text_height + line_spacing
        return wrapped_text, max_text_width, total_height

    wrap_width = 15  # Increased to account for smaller font
    wrapped_text, max_text_width, total_height = get_text_dimensions(title, font, wrap_width)

    while (max_text_width > max_width or total_height > max_height or len(wrapped_text) < 4) and font_size > 40:
        if len(wrapped_text) < 4:
            wrap_width -= 1
        font_size -= 5
        font = None
        for font_path in font_paths:
            try:
                font = ImageFont.truetype(font_path, font_size)
                break
            except:
                continue
        if not font:
            font = ImageFont.load_default()
        wrapped_text, max_text_width, total_height = get_text_dimensions(title, font, wrap_width)

    while total_height < min_height and font_size < 200 and len(wrapped_text) <= 5:
        font_size += 5
        font = None
        for font_path in font_paths:
            try:
                font = ImageFont.truetype(font_path, font_size)
                break
            except:
                continue
        if not font:
            font = ImageFont.load_default()
        wrapped_text, max_text_width, total_height = get_text_dimensions(title, font, wrap_width)

    text_area_height = total_height + 40
    text_area = Image.new("RGBA", (1080, text_area_height), (0, 0, 0, int(255 * 0.7)))
    text_draw = ImageDraw.Draw(text_area)

    current_y = 20
    for line in wrapped_text:
        text_bbox = text_draw.textbbox((0, 0), line, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_x = (1080 - text_width) // 2
        text_draw.text((text_x, current_y), line, font=font, fill=(255, 255, 255), stroke_width=3, stroke_fill=(0, 0, 0))
        current_y += (text_bbox[3] - text_bbox[1]) + line_spacing

    text_y = (1920 - text_area_height) // 2
    final_image = final_image.convert("RGBA")
    final_image.paste(text_area, (0, text_y), text_area)
    final_image = final_image.convert("RGB")

    final_image.save(output_path)
    print(f"  Saved title image at: {output_path}")

title_image_path = os.path.join(output_dir, "title_image.jpg")
create_title_image(title_text, bg_image_url, title_image_path)

# Stage 4: Download images
def download_images_with_icrawler(keyword, num_images, output_dir):
    print("Stage 4: Attempting to download images...")
    keyword_dir = os.path.join(output_dir, re.sub(r'[^\w\s-]', '', keyword.replace(' ', '_'))[:50])
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
            target_ratio = 1080 / 1920

            if img_ratio > target_ratio:
                new_height = 1920
                new_width = int(new_height * img_ratio)
            else:
                new_width = 1080
                new_height = int(new_width / img_ratio)

            img = img.resize((new_width, new_height), Image.LANCZOS)
            final_image = Image.new("RGB", (1080, 1920), (0, 0, 0))
            paste_x = (1080 - new_width) // 2
            paste_y = (1920 - new_height) // 2
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
additional_images = download_images_with_icrawler(keyword, 7, output_dir)
image_paths = [title_image_path] + additional_images
print(f"  Retrieved {len(additional_images)} images")

# Stage 5: Create video with varied transitions
def create_video(image_paths, audio_path, output_path):
    print("Stage 5: Creating video...")
    try:
        audio = AudioFileClip(audio_path)
        audio_duration = audio.duration
    except Exception as e:
        print(f"  Error loading audio {audio_path}: {e}. Exiting.")
        exit(1)

    clips = []
    duration_per_image = audio_duration / max(len(image_paths), 1)

    # Define transition effects
    def zoom_in(t, duration):
        return 1 + 0.02 * t  # Original zoom-in

    def pan_left(t, duration):
        return (1, 0.05 * t / duration)  # Move right (image shifts left)

    def pan_right(t, duration):
        return (1, -0.05 * t / duration)  # Move left (image shifts right)

    def pan_up(t, duration):
        return (0.05 * t / duration, 1)  # Move down (image shifts up)

    def pan_down(t, duration):
        return (-0.05 * t / duration, 1)  # Move up (image shifts down)

    transitions = [zoom_in, pan_left, pan_right, pan_up, pan_down]

    for img_path in image_paths:
        try:
            clip = ImageClip(img_path, duration=duration_per_image)
            # Randomly select a transition
            transition = random.choice(transitions)
            clip = clip.set_position(transition if transition != zoom_in else lambda t: ('center', 'center')).resize(transition if transition == zoom_in else lambda t: 1)
            clips.append(clip)
            print(f"  Applied transition {transition.__name__} to {img_path}")
        except Exception as e:
            print(f"  Warning: Failed to process image {img_path}: {e}")
            continue

    if not clips:
        print("  Error: No valid clips to create video. Exiting.")
        exit(1)

    try:
        video = concatenate_videoclips(clips, method="compose")
        video = video.set_audio(audio)
        video.write_videofile(output_path, codec="libx264", audio_codec="aac", fps=24, bitrate="1000k", audio_bitrate="64k", ffmpeg_params=["-preset", "ultrafast"])
        print(f"  Saved video at: {output_path}")
    except Exception as e:
        print(f"  Error saving video: {e}. Exiting.")
        exit(1)

output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
create_video(image_paths, audio_path, output_video_path)

print(f"Video created successfully at: {output_video_path}")

# Clean up temporary files (except video)
print("Cleaning up temporary files...")
for file in glob.glob(os.path.join(output_dir, keyword.replace(' ', '_')[:50], "*.jpg")):
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
