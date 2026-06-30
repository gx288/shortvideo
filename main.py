import os
os.environ["GRPC_ENABLE_FORK_SUPPORT"] = "1"
os.environ["GRPC_POLL_STRATEGY"] = "poll"

import re
import requests
import subprocess
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip
import glob
from google.cloud import texttospeech
import gspread
from google.oauth2.service_account import Credentials
import random
import unicodedata

# ==================== CẤU HÌNH ====================
NUM_VIDEOS_TO_CREATE = 1
WORKSHEET_LIST = ["LiveScience_Raw"]
Image.ANTIALIAS = Image.LANCZOS

# ==================== HÀM HỖ TRỢ ====================
def clean_filename(text, max_length=50):
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII')
    text = text.replace(' ', '_')
    text = re.sub(r'[^\w-]', '', text)
    text = re.sub(r'_+', '_', text)
    text = text[:max_length].strip('_')
    if not text or text == '':
        text = f"video_{random.randint(1000, 9999)}"
    return text.lower()

# ==================== MAIN SETUP ====================
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

print("Stage 0: Initializing Google Sheets...")
scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('google_sheets_key.json', scopes=scopes)
gc = gspread.authorize(creds)
SHEET_ID = '14tqKftTqlesnb0NqJZU-_f1EsWWywYqO36NiuDdmaTo'

def get_column_indices(worksheet):
    header = worksheet.row_values(1)
    cols = {
        "link": None,
        "status": None,
        "title": None,
        "content": None,
        "image": None
    }
    
    for idx, cell in enumerate(header, 1):
        c = str(cell).strip().lower()
        if c == "link video":
            cols["link"] = idx
        elif c == "đã đăng video?":
            cols["status"] = idx
        elif "tiêu đề" in c and "việt" in c:
            cols["title"] = idx
        elif "nội dung" in c and "việt" in c:
            cols["content"] = idx
        elif c in ["ảnh nền", "image url", "image", "ảnh"]:
            cols["image"] = idx
            
    # Fallback cho form cũ (Khoahocyhoc)
    if not cols["link"]: cols["link"] = 11
    if not cols["status"]: cols["status"] = 12
    if not cols["title"]: cols["title"] = 2
    if not cols["content"]: cols["content"] = 2
    if not cols["image"]: cols["image"] = 4
    
    return cols

videos_created = 0

for worksheet_name in WORKSHEET_LIST:
    if videos_created >= NUM_VIDEOS_TO_CREATE:
        break
    print(f"\nChecking worksheet: {worksheet_name}")
    try:
        worksheet = gc.open_by_key(SHEET_ID).worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Error: Worksheet '{worksheet_name}' not found. Skipping.")
        continue

    try:
        cols = get_column_indices(worksheet)
        LINK_COL = cols["link"]
        STATUS_COL = cols["status"]
        TITLE_COL = cols["title"]
        CONTENT_COL = cols["content"]
        IMAGE_COL = cols["image"]
        print(f"Tìm thấy các cột: Title={TITLE_COL}, Content={CONTENT_COL}, Image={IMAGE_COL}, Link={LINK_COL}, Status={STATUS_COL}")
    except ValueError as e:
        print(f"Lỗi tiêu đề cột trong sheet '{worksheet_name}': {e}")
        continue

    rows = worksheet.get_all_values()
    for i, row in enumerate(rows):
        if videos_created >= NUM_VIDEOS_TO_CREATE:
            break
        if i == 0: # Bỏ qua header
            continue

        link_value = row[LINK_COL - 1].strip() if len(row) >= LINK_COL else ""
        if link_value != "":
            continue

        print(f"Processing row {i + 1} in worksheet '{worksheet_name}'...")

        # ==================== LƯU THÔNG TIN DÒNG ====================
        selected_row_num = i + 1
        selected_row = row
        
        raw_title = selected_row[TITLE_COL - 1] if len(selected_row) >= TITLE_COL else ''
        raw_content = selected_row[CONTENT_COL - 1] if len(selected_row) >= CONTENT_COL else ''
        
        # Nếu dùng chung 1 cột (như form cũ Khoahocyhoc)
        if TITLE_COL == CONTENT_COL:
            raw_text = raw_title
            raw_text = re.sub(r'\*+', '', raw_text)
            raw_text = re.sub(r'[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF\U00002702-\U000027B0\U000024C2-\U0001F251]+', '', raw_text)
            raw_text = re.sub(r'#\w+\s*', '', raw_text)
            lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
            title_text = lines[0].replace('Tiêu đề:', '').strip() if lines else 'Untitled'
            content_text = '\n'.join(lines[1:]) if len(lines) > 1 else title_text
        else:
            # Nếu đã tách riêng cột Tiêu đề và Nội dung (LiveScience_Raw)
            title_text = re.sub(r'\*+', '', raw_title).strip()
            content_text = re.sub(r'\*+', '', raw_content).strip()
            
        clean_title = clean_filename(title_text)

        print(f"Original title: {title_text}")
        print(f"Clean title: {clean_title}")

        bg_image_url = selected_row[IMAGE_COL - 1] if len(selected_row) >= IMAGE_COL else 'https://via.placeholder.com/1080x1920?text=No+Image'

        # ==================== TTS ====================
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
            response = client.synthesize_speech(input=synthesis_input, voice=voice, audio_config=audio_config)
            audio_path = os.path.join(output_dir, "voiceover.mp3")
            with open(audio_path, "wb") as out:
                out.write(response.audio_content)
            print(f"Saved audio at: {audio_path}")
        except Exception as e:
            print(f"Error creating audio: {e}. Skipping.")
            continue

        # Cắt 55s
        try:
            temp_audio = os.path.join(output_dir, "temp_voiceover.mp3")
            subprocess.run([
                "ffmpeg", "-i", audio_path, "-t", "55", "-c:a", "mp3", "-b:a", "96k", temp_audio
            ], check=True, capture_output=True)
            os.replace(temp_audio, audio_path)
            print("Cut audio to 55s")
        except Exception as e:
            print(f"Warning: Failed to cut audio: {e}")

        # ==================== TITLE IMAGE ====================
        def create_title_image(title, bg_image_url, output_path):
            print("Stage 3: Creating title image...")
            try:
                bg_image = Image.open(requests.get(bg_image_url, stream=True, timeout=10).raw).convert("RGB")
            except Exception as e:
                print(f"Warning: Failed to download bg: {e}. Using black.")
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
            left = (new_width - target_size[0]) // 2
            top = (new_height - target_size[1]) // 2
            bg_image = bg_image.crop((left, top, left + target_size[0], top + target_size[1]))
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
                    print(f"Using font: {fp}")
                    break
                except:
                    continue
            if not font:
                print("Error: No font found.")
                return False

            max_width = 576
            min_height = 300
            max_height = 768
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
                for line in wrapped_text:
                    text_bbox = draw.textbbox((0, 0), line, font=font)
                    text_height = text_bbox[3] - text_bbox[1]
                    total_height += text_height + line_spacing
                return wrapped_text, total_height - line_spacing

            wrapped_text, total_height = get_text_dimensions(title, font, wrap_width)
            attempt = 0
            while (total_height < min_height or total_height > max_height) and wrap_width >= 10 and attempt < 20:
                wrap_width += -1 if total_height < min_height else 1
                wrapped_text, total_height = get_text_dimensions(title, font, wrap_width)
                attempt += 1

            text_area_height = total_height + 60
            text_area = Image.new("RGBA", (720, text_area_height), (0, 0, 0, int(255 * 0.7)))
            text_draw = ImageDraw.Draw(text_area)
            current_y = 30
            for line in wrapped_text:
                text_bbox = text_draw.textbbox((0, 0), line, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_x = (720 - text_width) // 2
                text_draw.text((text_x, current_y), line, font=font, fill=(255, 255, 255), stroke_width=2, stroke_fill=(0, 0, 0))
                current_y += (text_bbox[3] - text_bbox[1]) + line_spacing

            # Không dán text_area vào bg_image nữa để bg_image là nền trơn (phục vụ hiệu ứng panning)
            bg_image = bg_image.convert("RGB")
            bg_image.save(output_path)
            print(f"Saved title image: {output_path}")
            text_area.save(os.path.join(output_dir, "text_overlay.png"))
            return True

        title_image_path = os.path.join(output_dir, "title_image.jpg")
        if not create_title_image(title_text, bg_image_url, title_image_path):
            continue
        text_overlay_path = os.path.join(output_dir, "text_overlay.png")

        # ==================== TẢI ẢNH BING (ĐÃ TẮT ĐỂ CHỐNG 18+) ====================
        # Đã vô hiệu hóa theo yêu cầu để tránh rủi ro tải nhầm ảnh nhạy cảm
        image_paths = [title_image_path]
        print(f"Total images used: {len(image_paths)}")

        # ==================== TẠO VIDEO ====================
        def create_video(image_paths, audio_path, output_path):
            print("Stage 5: Creating video...")
            try:
                audio = AudioFileClip(audio_path)
                audio_duration = audio.duration
                print(f"Audio duration: {audio_duration:.2f}s")
            except Exception as e:
                print(f"Error loading audio: {e}")
                return False

            text_overlay = Image.open(text_overlay_path).convert("RGBA")
            clips = []
            durations = [7.0] + [2.0] * (len(image_paths) - 1)
            total_duration = sum(durations)

            while total_duration > audio_duration and len(durations) > 1:
                durations.pop()
                image_paths.pop()
                total_duration = sum(durations)

            if total_duration > audio_duration:
                durations[-1] = audio_duration - sum(durations[:-1])
                if durations[-1] < 0.5:
                    durations.pop()
                    image_paths.pop()

            for i, (img_path, duration) in enumerate(zip(image_paths, durations)):
                try:
                    img = Image.open(img_path).convert("RGB")
                    img_ratio = img.width / img.height
                    
                    # Kích thước khung hình lớn hơn 15% (828x1472) để có dư không gian cho Panning
                    target_ratio = 828 / 1472
                    if img_ratio > target_ratio:
                        new_height = 1472
                        new_width = int(new_height * img_ratio)
                    else:
                        new_width = 828
                        new_height = int(new_width / img_ratio)
                        
                    img = img.resize((new_width, new_height), Image.LANCZOS)
                    left = (new_width - 828) // 2
                    top = (new_height - 1472) // 2
                    img = img.crop((left, top, left + 828, top + 1472))
                    
                    np_img = np.array(img.convert("RGB"))
                    clip = ImageClip(np_img).set_duration(duration)
                    
                    # Chọn ngẫu nhiên hướng di chuyển của camera
                    effect = random.choice(['pan_left', 'pan_right', 'pan_up', 'pan_down'])
                    
                    # Chuẩn bị thông số để blend text overlay (text_overlay có kích thước 720xH)
                    text_np = np.array(text_overlay)
                    text_rgb = text_np[:, :, :3]
                    text_alpha = (text_np[:, :, 3:4] / 255.0)
                    
                    target_center_y = 1280 - (1280 // 3)
                    text_y = target_center_y - (text_overlay.height // 2)
                    text_y = max(0, text_y)
                    
                    def pan_and_overlay(get_frame, t):
                        frame = get_frame(t)
                        progress = t / duration
                        
                        if effect == 'pan_left':
                            x1 = int((1.0 - progress) * (828 - 720))
                            y1 = (1472 - 1280) // 2
                        elif effect == 'pan_right':
                            x1 = int(progress * (828 - 720))
                            y1 = (1472 - 1280) // 2
                        elif effect == 'pan_up':
                            x1 = (828 - 720) // 2
                            y1 = int((1.0 - progress) * (1472 - 1280))
                        else: # pan_down
                            x1 = (828 - 720) // 2
                            y1 = int(progress * (1472 - 1280))
                            
                        # Cắt khung hình 720x1280 từ ảnh nền to
                        cropped = frame[y1:y1+1280, x1:x1+720].copy()
                        
                        # Blend text overlay bằng numpy vào đúng tọa độ text_y
                        roi = cropped[text_y:text_y+text_overlay.height, :]
                        cropped[text_y:text_y+text_overlay.height, :] = roi * (1.0 - text_alpha) + text_rgb * text_alpha
                        
                        return cropped.astype(np.uint8)
                        
                    clip = clip.fl(lambda gf, t: pan_and_overlay(gf, t))
                    
                    clips.append(clip)
                    print(f"Image {i}: {duration:.1f}s - Effect: {effect}")
                except Exception as e:
                    print(f"Error processing image {img_path}: {e}")
                    continue

            if not clips:
                return False

            try:
                video = concatenate_videoclips(clips, method="compose")
                video = video.set_audio(audio)
                video.write_videofile(
                    output_path,
                    codec="libx265",
                    audio_codec="aac",
                    fps=15,
                    bitrate="300k",
                    audio_bitrate="96k",
                    ffmpeg_params=["-preset", "medium"]
                )
                print(f"Saved video: {output_path}")
                for f in glob.glob(os.path.join(output_dir, "temp_frame_*.png")):
                    try: os.remove(f)
                    except: pass
                return True
            except Exception as e:
                print(f"Error saving video: {e}")
                return False

        output_video_path = os.path.join(output_dir, f"output_video_{clean_title}.mp4")
        if create_video(image_paths, audio_path, output_video_path):
            videos_created += 1
            file_size_mb = os.path.getsize(output_video_path) / (1024 * 1024)
            print(f"Video created: {output_video_path}")
            print(f"Size: {file_size_mb:.2f} MB")
            
            # Cập nhật trực tiếp vào Google Sheets
            video_url = f"https://raw.githubusercontent.com/gx288/shortvideo/main/output/output_video_{clean_title}.mp4"
            try:
                worksheet.update_cell(selected_row_num, LINK_COL, video_url)
                if file_size_mb > 5:
                    worksheet.update_cell(selected_row_num, STATUS_COL, ">5MB")
                else:
                    worksheet.update_cell(selected_row_num, STATUS_COL, "")
                print(f"✅ Đã ghi link trực tiếp vào dòng {selected_row_num}, cột {LINK_COL}")
            except Exception as e:
                print(f"❌ Lỗi khi cập nhật sheet: {e}")
                
        else:
            print(f"Failed to create video for row {selected_row_num}")
            continue

        # Cleanup
        print("Cleaning up temporary files...")
        for temp_file in [audio_path, title_image_path, text_overlay_path]:
            if os.path.exists(temp_file):
                try: os.remove(temp_file)
                except: pass
        print("Cleanup complete.")

if videos_created == 0:
    print("No videos created. Exiting.")
    exit(1)

print(f"Successfully created {videos_created} video(s).")
