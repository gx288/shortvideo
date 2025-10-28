import os
import re
import requests
import subprocess
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip, CompositeVideoClip, TextClip
import glob
from google.cloud import texttospeech
import gspread
from google.oauth2.service_account import Credentials
import random
import unicodedata
import sys

# ====================== CONFIG ======================
NUM_VIDEOS_TO_CREATE = 1
WORKSHEET_LIST = ["Phòng mạch", "Sheet2", "Sheet3"]
# ====================================================

Image.ANTIALIAS = Image.LANCZOS, Image.LANCZOS

def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    return text.lower() if text else f"video_{random.randint(1000, 9999)}"

# HÀM MỚI: GHI TIÊU ĐỀ VÀO ẢNH
def add_title_to_image(image_path, title, output_path):
    try:
        img = Image.open(image_path).convert("RGB")
        draw = ImageDraw.Draw(img)
        w, h = img.size

        font_size = 70
        try:
            font = ImageFont.truetype("Arial-Bold", font_size)
        except:
            font = ImageFont.load_default()

        max_w = w * 0.8
        lines = []
        words = title.split()
        line = ""
        for word in words:
            test = line + (" " if line else "") + word
            if draw.textbbox((0,0), test, font=font)[2] <= max_w:
                line = test
            else:
                lines.append(line)
                line = word
        if line: lines.append(line)

        line_h = font.getbbox("A")[3] - font.getbbox("A")[1] + 10
        total_h = len(lines) * line_h + 40
        box_h = total_h
        box_w = w * 0.86
        box_y = (h - box_h) // 2

        box = Image.new("RGBA", (int(box_w), int(box_h)), (0, 0, 0, 153))
        img.paste(box, (int((w - box_w)//2), int(box_y)), box)

        y = box_y + 20
        for line in lines:
            bbox = draw.textbbox((0,0), line, font=font)
            text_w = bbox[2] - bbox[0]
            x = (w - text_w) // 2
            draw.text((x, y), line, font=font, fill=(255,255,255),
                      stroke_width=2, stroke_fill=(0,0,0))
            y += line_h

        img.save(output_path)
        return True
    except Exception as e:
        print(f"  ERROR adding title: {e}")
        return False

# ====================== MAIN ======================
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
        print(f"  ERROR opening worksheet: {e}")
        sys.exit(1)

    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break
        if i == 0 or len(row) <= 7 or row[7].strip():
            continue

        selected_row = row
        selected_row_num = i + 1
        print(f"\nProcessing row {selected_row_num}...")

        try:
            # Extract content
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

            # STAGE 1: DOWNLOAD IMAGES FIRST
            print("Stage 1: Downloading images...")
            image_paths = []
            cover_path = os.path.join(output_dir, "cover.jpg")
            try:
                response = requests.get(bg_image_url, timeout=10)
                response.raise_for_status()
                with open(cover_path, "wb") as f:
                    f.write(response.content)
                img = Image.open(cover_path).convert("RGB")
                img = img.resize((720, 1280), Image.LANCZOS)
                img.save(cover_path)
                image_paths.append(cover_path)
            except Exception as e:
                print(f"  ERROR downloading cover: {e}")
                sys.exit(1)

            try:
                from icrawler.builtin import GoogleImageCrawler
                keyword_clean = clean_filename(title_text[:50])
                keyword_dir = os.path.join(output_dir, keyword_clean)
                os.makedirs(keyword_dir, exist_ok=True)
                GoogleImageCrawler(storage={'root_dir': keyword_dir}).crawl(
                    keyword=title_text[:50], max_num=9, min_size=(500, 500)
                )
                for img_file in glob.glob(os.path.join(keyword_dir, "*.jpg"))[:9]:
                    try:
                        img = Image.open(img_file).convert("RGB")
                        img = img.resize((720, 1280), Image.LANCZOS)
                        img.save(img_file)
                        image_paths.append(img_file)
                    except:
                        continue
            except:
                print("  icrawler failed.")

            if len(image_paths) < 2:
                image_paths = [image_paths[0]] * 10
            print(f"  {len(image_paths)} images ready.")

            # STAGE 2: CREATE AUDIO
            print("Stage 2: Creating TTS audio...")
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
                with open(audio_path, "wb") as f:
                    f.write(response.audio_content)

                # CẮT AUDIO – CHỈ SỬA ĐOẠN NÀY
                temp = audio_path + ".tmp"
                try:
                    subprocess.run([
                        "ffmpeg", "-y", "-i", audio_path, "-t", "55",
                        "-c:a", "libmp3lame", "-b:a", "96k", "-ar", "44100", temp
                    ], check=True, capture_output=True)
                    os.replace(temp, audio_path)
                    print("  Audio cut to 55s")
                except subprocess.CalledProcessError as e:
                    print(f"  ERROR: ffmpeg failed: {e}")
                    if os.path.exists(temp):
                        os.remove(temp)
                    sys.exit(1)

            except Exception as e:
                print(f"  ERROR: TTS failed: {e}")
                sys.exit(1)

            # THÊM: GHI TIÊU ĐỀ VÀO TỪNG ẢNH
            print("Stage 2.5: Adding title to all images...")
            titled_images = []
            for idx, img_path in enumerate(image_paths):
                titled_path = os.path.join(output_dir, f"titled_{idx}.jpg")
                if add_title_to_image(img_path, title_text, titled_path):
                    titled_images.append(titled_path)
                else:
                    titled_images.append(img_path)

            # STAGE 3: CREATE VIDEO
            print("Stage 3: Creating video...")
            try:
                audio = AudioFileClip(audio_path)
                total_duration = min(audio.duration, 55)
                duration_per_clip = total_duration / len(titled_images)

                clips = []
                transitions = [
                    lambda t, d: 1 + 0.2 * (t/d),
                    lambda t, d: 1.2 - 0.2 * (t/d),
                    lambda t, d: (0.2 * (t/d) * 720, 'center'),
                    lambda t, d: (-0.2 * (t/d) * 720, 'center'),
                    lambda t, d: ('center', 0.2 * (t/d) * 1280),
                    lambda t, d: ('center', -0.2 * (t/d) * 1280),
                ]

                for idx, path in enumerate(titled_images):
                    clip = ImageClip(path).set_duration(duration_per_clip)
                    trans = transitions[idx % len(transitions)]
                    if idx < 2:
                        clip = clip.resize(lambda t: trans(t, duration_per_clip))
                    else:
                        clip = clip.set_position(lambda t: trans(t, duration_per_clip))
                    clips.append(clip)

                video = concatenate_videoclips(clips, method="compose")

                # TITLE FULL DURATION (giữ nguyên)
                txt = TextClip(
                    title_text,
                    fontsize=70, color='white', font='Arial-Bold',
                    stroke_color='black', stroke_width=2,
                    size=(576, None), method='caption', align='center'
                ).set_position('center').set_duration(total_duration)

                bg = TextClip("", color='black', size=(620, txt.h + 40)
                            ).set_opacity(0.6).set_position('center').set_duration(total_duration)

                final = CompositeVideoClip([video, bg, txt]).set_audio(audio)

                output_video = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
                final.write_videofile(
                    output_video, codec="libx265", audio_codec="aac",
                    fps=15, bitrate="700k", audio_bitrate="96k",
                    ffmpeg_params=["-preset", "medium"], threads=4
                )

                print(f"  SUCCESS: {output_video}")
                videos_created += 1

                try:
                    worksheet.update_cell(selected_row_num, 8, "DONE")
                except:
                    pass

            except Exception as e:
                print(f"  ERROR creating video: {e}")
                sys.exit(1)

            # Cleanup
            for f in [audio_path, cover_path] + titled_images:
                if os.path.exists(f): os.remove(f)
            import shutil
            for d in glob.glob(os.path.join(output_dir, clean_filename(title_text[:50]))):
                if os.path.isdir(d): shutil.rmtree(d, ignore_errors=True)

        except Exception as e:
            print(f"  FATAL ERROR: {e}")
            sys.exit(1)

if videos_created == 0:
    print("No videos created.")
    sys.exit(1)
else:
    print(f"\nDONE: {videos_created} video(s) created.")
    sys.exit(0)
