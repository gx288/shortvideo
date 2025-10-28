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

# Configuration
NUM_VIDEOS_TO_CREATE = 1
WORKSHEET_LIST = ["Phòng mạch", "Sheet2", "Sheet3"]

Image.ANTIALIAS = Image.LANCZOS

def check_ffmpeg():
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True)
        print("  ffmpeg is available")
        return True
    except:
        print("  ERROR: ffmpeg not found!")
        return False

def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    return text.lower() if text else f"video_{random.randint(1000, 9999)}"

# Directory setup
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)
print(f"Output directory: {output_dir}")

# Google Sheets
print("Initializing Google Sheets...")
try:
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
    gc = gspread.authorize(creds)
    SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'
except Exception as e:
    print(f"  ERROR: Google Sheets auth failed: {e}")
    sys.exit(1)

videos_created = 0

for worksheet_name in WORKSHEET_LIST:
    if videos_created >= NUM_VIDEOS_TO_CREATE:
        break
    print(f"\nChecking worksheet: {worksheet_name}")
    try:
        worksheet = gc.open_by_key(SHEET_ID).worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  Worksheet not found. Skipping.")
        continue
    except Exception as e:
        print(f"  ERROR: {e}")
        sys.exit(1)

    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break
        if i == 0: continue
        if len(row) > 7 and (not row[7] or row[7].strip() == ''):
            selected_row = row
            selected_row_num = i + 1
            print(f"\nProcessing row {selected_row_num}...")

            try:
                # Extract
                raw_content = re.sub(r'\*+', '', selected_row[1])
                raw_content = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_content)
                raw_content = re.sub(r'#\w+\s*', '', raw_content)
                lines = [line.strip() for line in raw_content.split('\n') if line.strip()]
                title_text = lines[0].replace('Tiêu đề:', '').strip() if lines else 'Untitled'
                content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text
                clean_title = clean_filename(title_text)

                with open(os.path.join(output_dir, "clean_title.txt"), "w", encoding="utf-8") as f:
                    f.write(clean_title)

                bg_image_url = selected_row[3] if len(selected_row) > 3 else 'https://via.placeholder.com/1080x1920'

                # === STAGE 1: DOWNLOAD IMAGES FIRST ===
                print("Stage 1: Downloading images...")
                keyword = title_text[:50]
                image_paths = []

                # Cover
                cover_path = os.path.join(output_dir, "cover.jpg")
                try:
                    response = requests.get(bg_image_url, timeout=10)
                    response.raise_for_status()
                    with open(cover_path, "wb") as f:
                        f.write(response.content)
                    img = Image.open(cover_path).convert("RGB")
                    img_ratio = img.width / img.height
                    target_ratio = 720 / 1280
                    if img_ratio > target_ratio:
                        new_h = 1280
                        new_w = int(new_h * img_ratio)
                    else:
                        new_w = 720
                        new_h = int(new_w / img_ratio)
                    img = img.resize((new_w, new_h), Image.LANCZOS)
                    final = Image.new("RGB", (720, 1280), (0, 0, 0))
                    final.paste(img, ((720 - new_w) // 2, (1280 - new_h) // 2))
                    final.save(cover_path)
                    image_paths.append(cover_path)
                except Exception as e:
                    print(f"  ERROR: {e}")
                    sys.exit(1)

                # Additional
                try:
                    from icrawler.builtin import GoogleImageCrawler
                    keyword_clean = clean_filename(keyword)
                    keyword_dir = os.path.join(output_dir, keyword_clean)
                    os.makedirs(keyword_dir, exist_ok=True)
                    GoogleImageCrawler(storage={'root_dir': keyword_dir}).crawl(keyword=keyword, max_num=10, min_size=(500, 500))
                    for img_file in glob.glob(os.path.join(keyword_dir, "*.jpg"))[:9]:
                        try:
                            img = Image.open(img_file).convert("RGB")
                            img_ratio = img.width / img.height
                            if img_ratio > target_ratio:
                                new_h = 1280
                                new_w = int(new_h * img_ratio)
                            else:
                                new_w = 720
                                new_h = int(new_w / img_ratio)
                            img = img.resize((new_w, new_h), Image.LANCZOS)
                            final = Image.new("RGB", (720, 1280), (0, 0, 0))
                            final.paste(img, ((720 - new_w) // 2, (1280 - new_h) // 2))
                            final.save(img_file)
                            image_paths.append(img_file)
                        except: continue
                except: pass

                if len(image_paths) < 2:
                    image_paths = [image_paths[0]] * 10

                # === STAGE 2: CREATE AUDIO (SAU IMAGES) ===
                print("Stage 2: Creating audio...")
                audio_path = os.path.join(output_dir, "voiceover.mp3")
                try:
                    client = texttospeech.TextToSpeechClient.from_service_account_file('google_tts_key.json')
                    response = client.synthesize_speech(
                        input=texttospeech.SynthesisInput(text=content_text),
                        voice=texttospeech.VoiceSelectionParams(language_code="vi-VN", name="vi-VN-Wavenet-C"),
                        audio_config=texttospeech.AudioConfig(
                            audio_encoding=texttospeech.AudioEncoding.MP3,
                            speaking_rate=1.25,
                            sample_rate_hertz=44100
                        )
                    )
                    with open(audio_path, "wb") as out:
                        out.write(response.audio_content)
                    print(f"  Saved audio: {audio_path}")

                    # CẮT AUDIO – CHỈ SỬA 1 DÒNG NÀY
                    temp_audio = os.path.join(output_dir, "temp_voiceover.mp3")
                    subprocess.run([
                        "ffmpeg", "-y", "-i", audio_path, "-t", "55",
                        "-c:a", "libmp3lame", "-ar", "44100", "-b:a", "96k", temp_audio
                    ], check=True, capture_output=True)
                    os.replace(temp_audio, audio_path)
                    print("  Audio cut to 55s")

                except Exception as e:
                    print(f"  ERROR: TTS/audio failed: {e}")
                    sys.exit(1)

                # === STAGE 3: VIDEO + TITLE FULL ===
                print("Stage 3: Creating video...")
                try:
                    audio = AudioFileClip(audio_path)
                    total_duration = min(audio.duration, 55)
                    duration_per_clip = total_duration / len(image_paths)

                    clips = []
                    transitions = [
                        lambda t, d: 1 + 0.2 * (t/d),
                        lambda t, d: 1.2 - 0.2 * (t/d),
                        lambda t, d: (0.2 * (t/d) * 720, 'center'),
                        lambda t, d: (-0.2 * (t/d) * 720, 'center'),
                        lambda t, d: ('center', 0.2 * (t/d) * 1280),
                        lambda t, d: ('center', -0.2 * (t/d) * 1280),
                    ]

                    for idx, path in enumerate(image_paths):
                        clip = ImageClip(path).set_duration(duration_per_clip)
                        trans = transitions[idx % len(transitions)]
                        if idx < 2:
                            clip = clip.resize(lambda t: trans(t, duration_per_clip))
                        else:
                            clip = clip.set_position(lambda t: trans(t, duration_per_clip))
                        clips.append(clip)

                    video = concatenate_videoclips(clips, method="compose")

                    # TITLE OVERLAY FULL VIDEO
                    txt_clip = TextClip(
                        title_text, fontsize=70, color='white', font='Arial-Bold',
                        stroke_color='black', stroke_width=2,
                        size=(576, None), method='caption', align='center'
                    ).set_position('center').set_duration(total_duration)

                    bg_txt = TextClip("", color='black', size=(620, txt_clip.h + 40)
                                    ).set_opacity(0.6).set_position('center').set_duration(total_duration)

                    final = CompositeVideoClip([video, bg_txt, txt_clip]).set_audio(audio)

                    output_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
                    final.write_videofile(
                        output_path, codec="libx265", audio_codec="aac",
                        fps=15, bitrate="700k", audio_bitrate="96k",
                        ffmpeg_params=["-preset", "medium"]
                    )
                    print(f"  SUCCESS: {output_path}")
                    videos_created += 1

                    try:
                        worksheet.update_cell(selected_row_num, 8, "DONE")
                    except: pass

                except Exception as e:
                    print(f"  ERROR: Video failed: {e}")
                    sys.exit(1)

                # Cleanup
                for f in [audio_path, cover_path]:
                    if os.path.exists(f): os.remove(f)
                import shutil
                if os.path.exists(os.path.join(output_dir, clean_filename(keyword))):
                    shutil.rmtree(os.path.join(output_dir, clean_filename(keyword)), ignore_errors=True)

            except Exception as e:
                print(f"  FATAL: {e}")
                sys.exit(1)

if videos_created == 0:
    print("No videos.")
    sys.exit(1)
else:
    print(f"DONE: {videos_created} video(s).")
    sys.exit(0)
