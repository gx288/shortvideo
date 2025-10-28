import os
import re
import requests
import subprocess
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip
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

def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    return text.lower() if text else f"video_{random.randint(1000, 9999)}"

def draw_text_on_image(image_path, text, output_path):
    try:
        img = Image.open(image_path).convert("RGB")
        draw = ImageDraw.Draw(img)
        width, height = img.size

        # Font
        font_size = 70
        try:
            font = ImageFont.truetype("Arial-Bold", font_size)
        except:
            try:
                font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", font_size)
            except:
                font = ImageFont.load_default()

        # Bọc text
        max_width = width * 0.8
        lines = []
        words = text.split()
        current = ""
        for word in words:
            test = current + (" " if current else "") + word
            bbox = draw.textbbox((0, 0), test, font=font)
            if bbox[2] - bbox[0] <= max_width:
                current = test
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)

        # Tính chiều cao
        line_height = font.getbbox("A")[3] - font.getbbox("A")[1] + 10
        total_h = len(lines) * line_height + 40
        box_h = total_h
        box_w = width * 0.86
        box_y = (height - box_h) // 2

        # Vẽ nền mờ
        box = Image.new("RGBA", (int(box_w), int(box_h)), (0, 0, 0, 153))
        img.paste(box, (int((width - box_w) // 2), int(box_y)), box)

        # Vẽ chữ
        y = box_y + 20
        for line in lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            w = bbox[2] - bbox[0]
            x = (width - w) // 2
            draw.text((x, y), line, font=font, fill=(255, 255, 255),
                      stroke_width=2, stroke_fill=(0, 0, 0))
            y += line_height

        img.save(output_path)
        return True
    except Exception as e:
        print(f"  ERROR drawing text on {image_path}: {e}")
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

            bg_image_url = selected_row[3] if len(selected_row) > 3 else 'https://via.placeholder.com/720x1280'

            # STAGE 1: DOWNLOAD IMAGES
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

            # Additional images
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
                print("  icrawler failed or not installed.")

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

                # Cut to 55s
                temp = audio_path + ".tmp"
                subprocess.run([
                    "ffmpeg", "-y", "-i", audio_path, "-t", "55", "-c:a", "libmp3lame", "-b:a", "96k", "-ar", "44100", temp
                ], check=True, capture_output=True)
                os.replace(temp, audio_path)
            except Exception as e:
                print(f"  ERROR: TTS failed: {e}")
                sys.exit(1)

            # STAGE 3: DRAW TITLE ON EACH IMAGE
            print("Stage 3: Drawing title on each image...")
            processed_images = []
            for idx, img_path in enumerate(image_paths):
                output_img = os.path.join(output_dir, f"img_with_text_{idx}.jpg")
                if draw_text_on_image(img_path, title_text, output_img):
                    processed_images.append(output_img)
                else:
                    processed_images.append(img_path)  # fallback

            # STAGE 4: CREATE VIDEO
            print("Stage 4: Creating video...")
            try:
                audio = AudioFileClip(audio_path)
                total_duration = min(audio.duration, 55)
                duration_per_clip = total_duration / len(processed_images)

                clips = []
                transitions = [
                    lambda t, d: 1 + 0.2 * (t/d),
                    lambda t, d: 1.2 - 0.2 * (t/d),
                    lambda t, d: (0.2 * (t/d) * 720, 'center'),
                    lambda t, d: (-0.2 * (t/d) * 720, 'center'),
                    lambda t, d: ('center', 0.2 * (t/d) * 1280),
                    lambda t, d: ('center', -0.2 * (t/d) * 1280),
                ]

                for idx, path in enumerate(processed_images):
                    clip = ImageClip(path).set_duration(duration_per_clip)
                    trans = transitions[idx % len(transitions)]
                    if idx < 2:
                        clip = clip.resize(lambda t: trans(t, duration_per_clip))
                    else:
                        clip = clip.set_position(lambda t: trans(t, duration_per_clip))
                    clips.append(clip)

                final_video = concatenate_videoclips(clips, method="compose").set_audio(audio)

                output_video = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
                final_video.write_videofile(
                    output_video, codec="libx265", audio_codec="aac",
                    fps=15, bitrate="700k", audio_bitrate="96k",
                    ffmpeg_params=["-preset", "medium"], threads=4
                )

                print(f"  SUCCESS: {output_video}")
                print(f"  Size: {os.path.getsize(output_video)/(1024*1024):.2f} MB")
                videos_created += 1

                # Update Google Sheet
                try:
                    worksheet.update_cell(selected_row_num, 8, "DONE")
                except:
                    pass

            except Exception as e:
                print(f"  ERROR creating video: {e}")
                sys.exit(1)

            # Cleanup
            for f in [audio_path, cover_path] + processed_images:
                if os.path.exists(f): os.remove(f)
            import shutil
            for d in glob.glob(os.path.join(output_dir, clean_filename(title_text[:50]))):
                if os.path.isdir(d): shutil.rmtree(d, ignore_errors=True)

        except Exception as e:
            print(f"  FATAL ERROR at row {selected_row_num}: {e}")
            sys.exit(1)

if videos_created == 0:
    print("No videos created.")
    sys.exit(1)
else:
    print(f"\nDONE: {videos_created} video(s) created successfully.")
    sys.exit(0)
