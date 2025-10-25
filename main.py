import os
import requests
import subprocess
import textwrap
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip
import numpy as np
import glob
from google.cloud import texttospeech

# Fallback for ANTIALIAS in Pillow
from PIL import Image
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

# Text content (Unicode)
title_text = "ĂN NHIỀU 15 SIÊU THỰC PHẨM NÀY, ĐƯỜNG RUỘT SẼ “CẢM ƠN“ BẠN"
content_text = """Đường ruột được ví như hệ thống rễ quan trọng của cơ thể, nơi hấp thụ các chất dinh dưỡng thiết yếu để duy trì mọi chức năng sống. Việc chăm sóc đường ruột đúng cách là chìa khóa vàng cho sức khỏe tổng thể, và những siêu thực phẩm chính là trợ thủ đắc lực giúp bạn đạt được điều đó.

Các loại siêu thực phẩm không chỉ cung cấp nguồn dưỡng chất phong phú mà còn hỗ trợ đường ruột xây dựng một hàng rào bảo vệ vững chắc. Điều này giúp ngăn chặn hiệu quả sự xâm nhập của vi khuẩn có hại, ký sinh trùng, nấm men và các độc tố, từ đó giảm thiểu nguy cơ mắc các vấn đề tiêu hóa và tăng cường khả năng miễn dịch.

Bằng cách bổ sung đều đặn các siêu thực phẩm này vào chế độ ăn uống hàng ngày, đường ruột sẽ được nuôi dưỡng từ sâu bên trong, hoạt động hiệu quả hơn. Đây là phương pháp tự nhiên để duy trì một hệ tiêu hóa khỏe mạnh, giúp cơ thể hấp thu tối đa dưỡng chất và tràn đầy năng lượng, mang lại cảm giác dễ chịu và một cuộc sống chất lượng hơn."""

# Stage 2: Create audio with Google Cloud TTS
print("Stage 2: Creating audio with Google Cloud TTS...")
try:
    client = texttospeech.TextToSpeechClient()
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

# Stage 3: Create title image (unchanged)
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
    font_size = 150
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

    wrap_width = 12
    wrapped_text, max_text_width, total_height = get_text_dimensions(title, font, wrap_width)

    while (max_text_width > max_width or total_height > max_height or len(wrapped_text) < 4) and font_size > 50:
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

    while total_height < min_height and font_size < 250 and len(wrapped_text) <= 5:
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
create_title_image(
    title_text,
    "https://cdn.24h.com.vn/upload/4-2025/images/2025-10-20/1760944625-616-thumbnail-width740height495_anh_cat_3_2_anh_cat_4_3.jpg",
    title_image_path
)

# Stage 4: Download images (unchanged)
def download_images_with_icrawler(keyword, num_images, output_dir):
    print("Stage 4: Attempting to download images...")
    keyword_dir = os.path.join(output_dir, keyword.replace(' ', '_'))
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

keyword = "siêu thực phẩm đường ruột"
additional_images = download_images_with_icrawler(keyword, 7, output_dir)
image_paths = [title_image_path] + additional_images
print(f"  Retrieved {len(additional_images)} images")

# Stage 5: Create video (unchanged except for bitrate, already set)
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

    for img_path in image_paths:
        try:
            clip = ImageClip(img_path, duration=duration_per_image)
            clip = clip.resize(lambda t: 1 + 0.02 * t)
            clips.append(clip)
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

output_video_path = os.path.join(output_dir, "output_video.mp4")
create_video(image_paths, audio_path, output_video_path)

print(f"Video created successfully at: {output_video_path}")

# Clean up temporary files (unchanged)
print("Cleaning up temporary files...")
for file in glob.glob(os.path.join(output_dir, "siêu_thực_phẩm_đường_ruột", "*.jpg")):
    try:
        os.remove(file)
    except:
        pass
print("Cleanup complete.")
