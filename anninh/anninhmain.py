import os
import re
import requests
import subprocess
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, AudioFileClip
from google.cloud import texttospeech
import gspread
from google.oauth2.service_account import Credentials
import random
import unicodedata

# ==================== CẤU HÌNH ====================
NUM_VIDEOS_TO_CREATE = 1
WORKSHEET_LIST = ["anninhhinhsu"]

# ==================== HÀM HỖ TRỢ ====================
def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    if not text:
        text = f"video_{random.randint(1000, 9999)}"
    return text.lower()


def create_title_image(title, bg_image_url, output_path, text_overlay_path, output_dir):
    """Tạo ảnh title với text overlay - Chỉ dùng 1 ảnh"""
    print("Stage 3: Creating title image...")
    try:
        bg_image = Image.open(requests.get(bg_image_url, stream=True, timeout=10).raw).convert("RGB")
    except Exception as e:
        print(f"Warning: Failed to download background: {e}. Using black background.")
        bg_image = Image.new("RGB", (720, 1280), (0, 0, 0))

    # Resize và crop về tỷ lệ 9:16 (720x1280)
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
    left = (new_width - target_size[0]) // 2
    top = (new_height - target_size[1]) // 2
    bg_image = bg_image.crop((left, top, left + target_size[0], top + target_size[1]))

    # Vẽ text
    draw = ImageDraw.Draw(bg_image)
    font_size = 48
    font_paths = [
        "Roboto-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ]
    font = None
    for fp in font_paths:
        try:
            font = ImageFont.truetype(fp, font_size)
            break
        except:
            continue

    if not font:
        print("Error: No font found.")
        return False

    # Wrap text
    max_width = 576
    line_spacing = 15
    wrapped_text = []
    current_line = ""
    for word in title.split():
        test_line = current_line + (" " if current_line else "") + word
        if draw.textbbox((0, 0), test_line, font=font)[2] <= max_width:
            current_line = test_line
        else:
            if current_line:
                wrapped_text.append(current_line)
            current_line = word
    if current_line:
        wrapped_text.append(current_line)

    # Tính chiều cao text
    total_height = 0
    for line in wrapped_text:
        bbox = draw.textbbox((0, 0), line, font=font)
        total_height += (bbox[3] - bbox[1]) + line_spacing
    total_height -= line_spacing

    text_area_height = total_height + 80
    text_area = Image.new("RGBA", (720, text_area_height), (0, 0, 0, int(255 * 0.7)))
    text_draw = ImageDraw.Draw(text_area)

    current_y = 30
    for line in wrapped_text:
        bbox = text_draw.textbbox((0, 0), line, font=font)
        text_width = bbox[2] - bbox[0]
        text_x = (720 - text_width) // 2
        text_draw.text((text_x, current_y), line, font=font, 
                      fill=(255, 255, 255), stroke_width=2, stroke_fill=(0, 0, 0))
        current_y += (bbox[3] - bbox[1]) + line_spacing

    # Dán text vào vị trí dưới
    target_center_y = 1280 - (1280 // 3)
    text_y = target_center_y - (text_area_height // 2)
    text_y = max(0, text_y)

    bg_image = bg_image.convert("RGBA")
    bg_image.paste(text_area, (0, text_y), text_area)
    bg_image = bg_image.convert("RGB")
    bg_image.save(output_path)
    text_area.save(text_overlay_path)

    print(f"Saved title image: {output_path}")
    return True


def create_video(title_image_path, audio_path, output_path, output_dir):
    """Tạo video chỉ dùng DUY NHẤT 1 ảnh lặp lại từ đầu đến cuối"""
    print("Stage 5: Creating video (lặp 1 ảnh duy nhất)...")
    try:
        audio = AudioFileClip(audio_path)
        audio_duration = audio.duration
        print(f"Audio duration: {audio_duration:.2f}s")
    except Exception as e:
        print(f"Error loading audio: {e}")
        return False

    # Tạo clip từ 1 ảnh và set duration = độ dài audio
    try:
        clip = ImageClip(title_image_path, duration=audio_duration)
        clip = clip.set_audio(audio)

        clip.write_videofile(
            output_path,
            codec="libx265",
            audio_codec="aac",
            fps=15,
            bitrate="300k",
            audio_bitrate="96k",
            ffmpeg_params=["-preset", "medium"],
            logger=None  # Tắt log moviepy để sạch hơn
        )
        print(f"Video saved successfully: {output_path}")
        print(f"Size: {os.path.getsize(output_path) / (1024 * 1024):.2f} MB")
        return True
    except Exception as e:
        print(f"Error creating video: {e}")
        return False


# ==================== MAIN SETUP ====================
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

print("Initializing Google Sheets...")
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'


def get_column_indices(worksheet):
    header = worksheet.row_values(1)
    link_col = status_col = None
    for idx, cell in enumerate(header, 1):
        if cell.strip() == "Link video":
            link_col = idx
        elif cell.strip() == "Đã đăng video?":
            status_col = idx
    if not link_col:
        raise ValueError("Không tìm thấy cột 'Link video'")
    if not status_col:
        raise ValueError("Không tìm thấy cột 'Đã đăng video?'")
    return {"link": link_col, "status": status_col}


videos_created = 0

for worksheet_name in WORKSHEET_LIST:
    if videos_created >= NUM_VIDEOS_TO_CREATE:
        break

    print(f"\nChecking worksheet: {worksheet_name}")
    try:
        worksheet = gc.open_by_key(SHEET_ID).worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Worksheet '{worksheet_name}' not found. Skipping.")
        continue

    try:
        cols = get_column_indices(worksheet)
        LINK_COL = cols["link"]
        STATUS_COL = cols["status"]
    except ValueError as e:
        print(f"Column error in sheet '{worksheet_name}': {e}")
        continue

    rows = worksheet.get_all_values()

    for i, row in enumerate(rows):
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break
        if i == 0:  # Skip header
            continue

        # Chỉ xử lý nếu chưa có link video
        link_value = row[LINK_COL - 1].strip() if len(row) >= LINK_COL else ""
        if link_value != "":
            continue

        print(f"\nProcessing row {i + 1} in worksheet '{worksheet_name}'...")

        # Ghi sheet hiện tại
        with open(os.path.join(output_dir, "current_sheet.txt"), "w", encoding="utf-8") as f:
            f.write(worksheet_name)

        # ==================== LẤY DỮ LIỆU TỪ ROW HIỆN TẠI ====================
        raw_content = row[1] if len(row) > 1 else ''
        raw_content = re.sub(r'\*+', '', raw_content)
        raw_content = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_content)
        raw_content = re.sub(r'#\w+\s*', '', raw_content)

        lines = [line.strip() for line in raw_content.split('\n') if line.strip()]

        if not lines:
            print("Dòng trống hoặc không có nội dung, bỏ qua.")
            continue

        title_text = lines[0].replace('Tiêu đề:', '').strip()
        content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text

        if not title_text.strip():
            print("Tiêu đề rỗng, bỏ qua dòng này.")
            continue

        clean_title = clean_filename(title_text)
        bg_image_url = row[3] if len(row) > 3 else 'https://via.placeholder.com/1080x1920?text=No+Image'

        print(f"Title: {title_text}")
        print(f"Content length: {len(content_text)} chars")

        # ==================== TẠO AUDIO ====================
        print("Stage 2: Creating audio with Google Cloud TTS...")
        audio_path = os.path.join(output_dir, "voiceover.mp3")
        try:
            client = texttospeech.TextToSpeechClient.from_service_account_file('google_tts_key.json')
            synthesis_input = texttospeech.SynthesisInput(text=content_text)
            voice = texttospeech.VoiceSelectionParams(language_code="vi-VN", name="vi-VN-Wavenet-C")
            audio_config = texttospeech.AudioConfig(
                audio_encoding=texttospeech.AudioEncoding.MP3,
                speaking_rate=1.25,
                pitch=0.0,
                sample_rate_hertz=44100
            )
            response = client.synthesize_speech(input=synthesis_input, voice=voice, audio_config=audio_config)

            with open(audio_path, "wb") as out:
                out.write(response.audio_content)

            # Cắt tối đa 55 giây
            temp_audio = os.path.join(output_dir, "temp_voiceover.mp3")
            subprocess.run([
                "ffmpeg", "-i", audio_path, "-t", "55", "-c:a", "mp3", "-b:a", "96k", temp_audio
            ], check=True, capture_output=True)
            os.replace(temp_audio, audio_path)
            print("Audio created and cut to 55s")

        except Exception as e:
            print(f"Error creating audio: {e}")
            continue

        # ==================== TẠO TITLE IMAGE ====================
        title_image_path = os.path.join(output_dir, "title_image.jpg")
        text_overlay_path = os.path.join(output_dir, "text_overlay.png")

        if not create_title_image(title_text, bg_image_url, title_image_path, text_overlay_path, output_dir):
            print("Failed to create title image.")
            continue

        # ==================== TẠO VIDEO ====================
        output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")

        if create_video(title_image_path, audio_path, output_video_path, output_dir):
            videos_created += 1
            print(f"✅ Video created successfully: {output_video_path}")
        else:
            print("Failed to create video.")
            continue

        # ==================== CLEANUP ====================
        print("Cleaning up temporary files...")
        for temp_file in [audio_path, title_image_path, text_overlay_path]:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass
        print("Cleanup completed.")

if videos_created == 0:
    print("No videos were created.")
    exit(1)

print(f"\n🎉 Successfully created {videos_created} video(s).")
