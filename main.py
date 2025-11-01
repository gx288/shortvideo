import os
import re
import requests
import subprocess
import textwrap
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip, CompositeVideoClip
import numpy as np
import glob
from google.cloud import texttospeech
import gspread
from google.oauth2.service_account import Credentials
import random
import unicodedata

# ==============================
# CẤU HÌNH
# ==============================
NUM_VIDEOS_TO_CREATE = 1
WORKSHEET_LIST = [
    "Phòng mạch",
    "Sheet2",
    "Sheet3",
]

# Fallback cho ANTIALIAS
Image.ANTIALIAS = Image.LANCZOS

# ==============================
# KIỂM TRA FFMPEG
# ==============================
def check_ffmpeg():
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, check=True)
        print("ffmpeg is available")
        return True
    except:
        print("Warning: ffmpeg not found!")
        return False

# ==============================
# LÀM SẠCH TÊN FILE
# ==============================
def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    return text.lower() or f"video_{random.randint(1000, 9999)}"

# ==============================
# RESIZE + CROP GIỮ NỀN TRONG
# ==============================
def resize_and_crop_transparent(img, target_size):
    target_width, target_height = target_size
    img_ratio = img.width / img.height
    target_ratio = target_width / target_height

    if img_ratio > target_ratio:
        new_height = target_height
        new_width = int(new_height * img_ratio)
    else:
        new_width = target_width
        new_height = int(new_width / img_ratio)

    img = img.resize((new_width, new_height), Image.LANCZOS)
    left = (new_width - target_width) // 2
    top = (new_height - target_height) // 2
    return img.crop((left, top, left + target_width, top + target_height))

# ==============================
# TẠO ẢNH TIÊU ĐỀ
# ==============================
def create_title_image(title, bg_image_url, output_path):
    print("Stage 3: Creating title image...")
    try:
        bg_image = Image.open(requests.get(bg_image_url, stream=True, timeout=10).raw).convert("RGB")
    except:
        print("Warning: Using black background")
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
    ]
    font = None
    for fp in font_paths:
        try:
            font = ImageFont.truetype(fp, font_size)
            print(f"Using font: {fp}")
            break
        except:
            continue
    if not font:
        font = ImageFont.load_default()
        print("Warning: Using default font")

    max_width = 576
    wrap_width = 30
    line_spacing = 15

    def get_text_size(text_lines):
        total_h = 0
        max_w = 0
        for line in text_lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
            max_w = max(max_w, w)
            total_h += h + line_spacing
        return max_w, total_h - line_spacing

    # Wrap text
    words = title.split()
    lines = []
    current = ""
    for word in words:
        test = current + (" " if current else "") + word
        if draw.textbbox((0, 0), test, font=font)[2] <= max_width:
            current = test
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)

    # Adjust if too tall
    while len(lines) > 5 and wrap_width > 15:
        wrap_width -= 2
        lines = textwrap.wrap(title, width=wrap_width)
    lines = lines[:5]

    _, total_h = get_text_size(lines)
    text_area_h = total_h + 60
    text_area = Image.new("RGBA", (720, text_area_h), (0, 0, 0, int(255 * 0.7)))
    text_draw = ImageDraw.Draw(text_area)
    y = 30
    for line in lines:
        bbox = text_draw.textbbox((0, 0), line, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        x = (720 - w) // 2
        text_draw.text((x, y), line, font=font, fill=(255, 255, 255), stroke_width=2, stroke_fill=(0, 0, 0))
        y += h + line_spacing

    text_y = (1280 - text_area_h) // 2
    final_image = final_image.convert("RGBA")
    final_image.paste(text_area, (0, text_y), text_area)
    final_image.convert("RGB").save(output_path, "JPEG", quality=95)
    print(f"Saved title image: {output_path}")

# ==============================
# TẢI CLIPART TỪ GOOGLE
# ==============================
def download_images_with_icrawler(keyword, num_images, output_dir):
    print("Stage 4: Downloading CLIPART images...")
    keyword_clean = clean_filename(keyword)
    keyword_dir = os.path.join(output_dir, keyword_clean)
    os.makedirs(keyword_dir, exist_ok=True)

    try:
        from icrawler.builtin import GoogleImageCrawler

        filters = dict(
            type='clipart',           # Chỉ lấy clipart
            size='large',
            color='color',
            license='commercial,other',
        )

        crawler = GoogleImageCrawler(storage={'root_dir': keyword_dir}, downloader_threads=5)
        search_kw = f"{keyword} clipart OR illustration OR vector OR icon OR drawing -photo -realistic"

        crawler.crawl(
            keyword=search_kw,
            filters=filters,
            max_num=num_images,
            min_size=(600, 600),
            file_idx_offset=0
        )
        print(f"Crawled clipart for: {search_kw}")

    except ImportError:
        print("Warning: icrawler not installed")
        return [title_image_path] * max(2, num_images)
    except Exception as e:
        print(f"Warning: Crawling failed: {e}")
        return [title_image_path] * max(2, num_images)

    # Xử lý ảnh
    files = glob.glob(os.path.join(keyword_dir, "*.*"))
    image_paths = []
    for path in files[:num_images]:
        try:
            img = Image.open(path).convert("RGBA")
            img = resize_and_crop_transparent(img, (720, 1280))
            out_path = path.rsplit(".", 1)[0] + "_clipart.png"
            img.save(out_path, "PNG")
            image_paths.append(out_path)
        except Exception as e:
            print(f"Failed to process: {e}")
    return image_paths or [title_image_path] * 2

# ==============================
# TẠO CLIP DUAL-LAYER (3/4 + BLUR ZOOM)
# ==============================
def create_dual_layer_clip(image_path, duration):
    try:
        img = Image.open(image_path).convert("RGB")
        base = resize_and_crop_transparent(img.convert("RGBA"), (720, 1280)).convert("RGB")

        # Background: blur + zoom
        bg = base.copy().filter(ImageFilter.GaussianBlur(20))
        bg_clip = ImageClip(np.array(bg), duration=duration)
        bg_clip = bg_clip.resize(lambda t: 1.1 + 0.1 * (t / duration))

        # Foreground: 3/4 screen
        fg_w = int(720 * 0.75)
        fg_h = int(fg_w * base.height / base.width)
        if fg_h > 1280 * 0.75:
            fg_h = int(1280 * 0.75)
            fg_w = int(fg_h * base.width / base.height)
        fg = base.resize((fg_w, fg_h), Image.LANCZOS)
        fg_clip = ImageClip(np.array(fg), duration=duration)
        fg_clip = fg_clip.set_position(('center', 'center'))

        return CompositeVideoClip([bg_clip, fg_clip])
    except:
        return ImageClip(image_path, duration=duration).resize(height=1280)

# ==============================
# TẠO VIDEO
# ==============================
def create_video(image_paths, audio_path, output_path):
    print("Stage 5: Creating video with dual-layer clipart effect...")
    try:
        audio = AudioFileClip(audio_path)
        audio_duration = min(audio.duration, 55)
    except Exception as e:
        print(f"Error loading audio: {e}")
        return False

    clips = []
    title_dur = min(5.0, audio_duration * 0.3)
    other_dur = (audio_duration - title_dur) / max(1, len(image_paths) - 1)

    # Title: full screen
    try:
        title_clip = ImageClip(image_paths[0], duration=title_dur).resize(height=1280)
        clips.append(title_clip)
    except:
        pass

    # Other clips: dual-layer
    for i in range(1, len(image_paths)):
        clip = create_dual_layer_clip(image_paths[i], other_dur)
        clip = clip.resize(lambda t: 1 + 0.03 * np.sin(t * np.pi / other_dur))
        clips.append(clip)

    if not clips:
        return False

    try:
        video = concatenate_videoclips(clips, method="compose")
        video = video.set_audio(audio)
        video.write_videofile(
            output_path,
            codec="libx265",
            audio_codec="aac",
            fps=24,
            bitrate="600k",
            audio_bitrate="128k",
            threads=4,
            preset="fast"
        )
        print(f"Video saved: {output_path}")
        return True
    except Exception as e:
        print(f"Error saving video: {e}")
        return False

# ==============================
# MAIN
# ==============================
if __name__ == "__main__":
    check_ffmpeg()
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    print(f"Created: {output_dir}")

    # Google Sheets
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
    gc = gspread.authorize(creds)
    SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'

    videos_created = 0

    for worksheet_name in WORKSHEET_LIST:
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break

        print(f"\nChecking worksheet: {worksheet_name}")
        try:
            worksheet = gc.open_by_key(SHEET_ID).worksheet(worksheet_name)
        except:
            print(f"Worksheet not found: {worksheet_name}")
            continue

        rows = worksheet.get_all_values()
        for i, row in enumerate(rows):
            if videos_created >= NUM_VIDEOS_TO_CREATE:
                break
            if i == 0 or len(row) <= 7 or row[7].strip():
                continue

            print(f"Processing row {i + 1}...")
            selected_row = row
            selected_row_num = i + 1

            # Nội dung
            raw = selected_row[1] if len(selected_row) > 1 else ''
            raw = re.sub(r'\*+', '', raw)
            raw = re.sub(r'[\U0001F600-\U0001F64F].*?', '', raw, flags=re.UNICODE)
            raw = re.sub(r'#\w+\s*', '', raw)
            lines = [l.strip() for l in raw.split('\n') if l.strip()]
            title_text = lines[0].replace('Tiêu đề:', '', 1).strip() if lines else 'Untitled'
            content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text

            clean_title = clean_filename(title_text)
            with open(os.path.join(output_dir, "clean_title.txt"), "w") as f:
                f.write(clean_title)

            bg_image_url = selected_row[3] if len(selected_row) > 3 else 'https://via.placeholder.com/1080x1920'
            print(f"Title: {title_text}")

            # TTS
            print("Stage 2: Creating audio...")
            try:
                client = texttospeech.TextToSpeechClient.from_service_account_file('google_tts_key.json')
                input_text = texttospeech.SynthesisInput(text=content_text)
                voice = texttospeech.VoiceSelectionParams(language_code="vi-VN", name="vi-VN-Wavenet-C")
                audio_config = texttospeech.AudioConfig(
                    audio_encoding=texttospeech.AudioEncoding.MP3,
                    speaking_rate=1.25,
                    sample_rate_hertz=44100
                )
                response = client.synthesize_speech(input=input_text, voice=voice, audio_config=audio_config)
                audio_path = os.path.join(output_dir, "voiceover.mp3")
                with open(audio_path, "wb") as f:
                    f.write(response.audio_content)
                print(f"Audio saved: {audio_path}")

                # Cắt 55s
                temp = audio_path.replace(".mp3", "_temp.mp3")
                subprocess.run([
                    "ffmpeg", "-i", audio_path, "-t", "55", "-c:a", "mp3", "-b:a", "96k", temp
                ], check=True, capture_output=True)
                os.replace(temp, audio_path)
            except Exception as e:
                print(f"Audio error: {e}")
                continue

            # Title image
            title_image_path = os.path.join(output_dir, "title_image.jpg")
            create_title_image(title_text, bg_image_url, title_image_path)
            if not os.path.exists(title_image_path):
                continue

            # Download clipart
            keyword = title_text[:50]
            additional_images = download_images_with_icrawler(keyword, 10, output_dir)
            image_paths = [title_image_path] + additional_images

            # Create video
            output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
            if create_video(image_paths, audio_path, output_video_path):
                videos_created += 1
                print(f"Video {videos_created} created: {output_video_path}")

            # Cleanup
            for f in glob.glob(os.path.join(output_dir, clean_filename(keyword), "*")):
                try: os.remove(f)
                except: pass
            for tmp in [audio_path, title_image_path]:
                if os.path.exists(tmp):
                    try: os.remove(tmp)
                    except: pass

    if videos_created == 0:
        print("No videos created.")
        exit(1)
    print(f"Done! Created {videos_created} video(s).")
