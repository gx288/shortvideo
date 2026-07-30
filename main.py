"""
main.py
=======
Main Entrypoint cho GitHub Actions Workflow và Local Runner.
Tự động tạo Video Short Drama Kịch Tính (< 3 phút):
- Nguồn truyện: 3,078 bài báo Tâm sự gia đình Afamily.vn (afamily_scraper/afamily_links.json)
- Nguồn video nền: 15,427 clip 9:16 DIY & Handmade (instagram/link_pool.json)
- Viết lại kịch bản AI kịch tính 3 Hồi (Hook 3s mở đầu)
- Tốc độ giọng đọc 1.2x dồn dập
- Nhạc nền nhacnen.mp3 (7% volume)
- Xuất chuẩn MP4 H.264 Baseline 720p yuv420p Faststart (2Mbps, xem được 100% mọi thiết bị)
"""

import os
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Call create_full_dramatic_video.py engine
from create_full_dramatic_video import create_guaranteed_video

if __name__ == "__main__":
    print("========================================")
    print("🚀 [DỰ ÁN SHORT VIDEO MỚI] BẮT ĐẦU TẠO VIDEO DRAMA HÃNG")
    print("========================================")
    create_guaranteed_video()
