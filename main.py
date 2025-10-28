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

# ====================== CONFIG ======================
NUM_VIDEOS_TO_CREATE = 1
WORKSHEET_LIST = [
    "Phòng mạch",
    "Sheet2",
    "Sheet3",
]
# ====================================================

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

# Google Sheets setup
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

# ====================== MAIN LOOP ======================
for worksheet_name in WORKSHEET_LIST:
    if videos_created >= NUM_VIDEOS_TO_CREATE:
        break

    print(f"\nChecking worksheet: {worksheet_name}")
    try:
        worksheet = gc.open_by_key(SHEET_ID).worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"  Worksheet '{worksheet_name}' not found. Skipping.")
        continue
    except Exception as e:
        print(f"  ERROR opening worksheet: {e}")
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

                with open(os.path.join(output_dir, "clean_title.txt"), "w", encoding="utf-8") as f:
                    f.write(clean_title)

                bg_image_url = selected_row[3] if len(selected_row) > 3 else 'https://via.placeholder.com/1080x1920?text=No+Image'
                print(f"  Cover image URL: {bg_image_url}")

                # === STAGE 1: DOWNLOAD IMAGES FIRST ===
                print("Stage 1: Downloading images...")
                keyword = title_text[:50]
                image_paths = []

                # Download cover image
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
                    print(f"  Cover downloaded: {cover_path}")
                except Exception as e:
                    print(f"  ERROR: Failed to download cover: {e}")
                    sys.exit(1)

                # Download additional images
                try:
                    from icrawler.builtin import GoogleImageCrawler
                    keyword_clean = clean_filename(keyword)
                    keyword_dir = os.path.join(output_dir, keyword_clean)
                    os.makedirs(keyword_dir, exist_ok=True)
                    GoogleImageCrawler(storage={'root_dir': keyword_dir}).crawl(
                        keyword=keyword, max_num=10, min_size=(500, 500)
                    )
                    print(f"  Crawled images for: {keyword}")
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
                        except Exception as e:
                            print(f"  Warning: Failed to process {img_file}: {e}")
                except ImportError:
                    print("  icrawler not installed. Using cover only.")
                except Exception as e:
                    print(f"  Crawling failed: {e}. Using cover only.")

                if len(image_paths) < 2:
                    image_paths = [image_paths[0]] * 10
                print(f"  Total images: {len(image_paths)}")

                # === STAGE 2: CREATE AUDIO (AFTER IMAGES) – ĐÃ SỬA LỖI 234 ===
                print("Stage 2: Creating audio with Google Cloud TTS...")
                audio_path = os.path.join(output_dir, "voiceover.mp3")
                temp_path = audio_path + ".tmp"
                cut_temp = audio_path + ".cut.tmp"
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

                    # GHI FILE AN TOÀN
                    with open(temp_path, "wb") as out:
                        out.write(response.audio_content)
                    print(f"  Raw audio saved: {temp_path}")

                    # KIỂM TRA FILE TRƯỚC KHI CẮT
                    probe = subprocess.run([
                        "ffprobe", "-v", "error", "-show_entries", "format=duration",
                        "-of", "default=noprint_wrappers=1:nokey=1", temp_path
                    ], capture_output=True, text=True)

                    if probe.returncode != 0:
                        print("  ERROR: MP3 file corrupted! ffprobe failed.")
                        raise Exception("Invalid MP3 from TTS")

                    duration = float(probe.stdout.strip())
                    print(f"  Audio duration: {duration:.2f}s")

                    # CẮT = 55s
                    cut_duration = min(duration, 55)
                    subprocess.run([
                        "ffmpeg", "-y", "-i", temp_path,
                        "-t", str(cut_duration),
                        "-c:a", "libmp3lame", "-b:a", "96k", "-ar", "44100",
                        cut_temp
                    ], check=True, capture_output=True)

                    # CHỈ THAY THẾ KHI THÀNH CÔNG
                    if os.path.exists(audio_path):
                        os.remove(audio_path)
                    os.replace(cut_temp, audio_path)
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

                    print(f"  Audio cut to {cut_duration:.1f}s: {audio_path}")

                except Exception as e:
                    print(f"  ERROR creating/cutting audio: {e}")
                    for p in [audio_path, temp_path, cut_temp]:
                        if 'p' in locals() and os.path.exists(p):
                            try: os.remove(p)
                            except: pass
                    sys.exit(1)

                # === STAGE 3: CREATE VIDEO WITH TITLE FULL DURATION ===
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
                            clip = clip.resize(trans)
                        else:
                            clip = clip.set_position(trans)
                        clips.append(clip)

                    video = concatenate_videoclips(clips, method="compose")

                    # TITLE OVERLAY FULL VIDEO
                    txt_clip = TextClip(
                        title_text,
                        fontsize=70, color='white', font='Arial-Bold',
                        stroke_color='black', stroke_width=2,
                        size=(576, None), method='caption', align='center'
                    ).set_position('center').set_duration(total_duration)

                    bg_txt = TextClip("", color='black', size=(620, txt_clip.h + 40)
                                    ).set_opacity(0.6).set_position('center').set_duration(total_duration)

                    final_video = CompositeVideoClip([video, bg_txt, txt_clip]).set_audio(audio)

                    output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
                    final_video.write_videofile(
                        output_video_path,
                        codec="libx265", audio_codec="aac", fps=15,
                        bitrate="700k", audio_bitrate="96k",
                        ffmpeg_params=["-preset", "medium"], threads=4
                    )
                    print(f"  SUCCESS: {output_video_path}")
                    print(f"  Size: {os.path.getsize(output_video_path)/(1024*1024):.2f} MB")

                    videos_created += 1

                    # UPDATE SHEET
                    try:
                        worksheet.update_cell(selected_row_num, 8, "DONE")
                    except:
                        pass

                except Exception as e:
                    print(f"  ERROR creating video: {e}")
                    sys.exit(1)

                # === CLEANUP ===
                print("Cleaning up...")
                for f in [audio_path, cover_path]:
                    if os.path.exists(f): os.remove(f)
                import shutil
                keyword_dir = os.path.join(output_dir, clean_filename(keyword))
                if os.path.exists(keyword_dir):
                    shutil.rmtree(keyword_dir, ignore_errors=True)
                print("Cleanup done.\n")

            except Exception as e:
                print(f"  FATAL ERROR at row {selected_row_num}: {e}")
                sys.exit(1)

# ====================== FINAL ======================
if videos_created == 0:
    print("No videos created.")
    sys.exit(1)
else:
    print(f"\nDONE: {videos_created} video(s) created successfully.")
    sys.exit(0)
