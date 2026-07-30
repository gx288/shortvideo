"""
instagram/batch_video_crawler.py
=================================
Thu thập 10,000 - 20,000 link video nền dọc (9:16 portrait) cho hashtag #diy và #handmade.
Chạy bằng Multithreading 10 luồng song song, chỉ lưu URL metadata (siêu nhẹ, không tốn ổ đĩa).
"""

import os
import sys
import json
import time
import subprocess
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

POOL_FILE  = os.path.join("instagram", "link_pool.json")
STATS_FILE = os.path.join("instagram", "link_pool_stats.json")

# 30 truy vấn hashtag DIY / Handmade mở rộng để đạt 10,000 - 20,000 video
HASHTAG_QUERIES = [
    "ytsearch1000:#shorts #diy",
    "ytsearch1000:#shorts #handmade",
    "ytsearch1000:#shorts #diycrafts",
    "ytsearch1000:#shorts #handcraft",
    "ytsearch1000:#shorts #crafts",
    "ytsearch1000:#shorts #doityourself",
    "ytsearch1000:#shorts #diyideas",
    "ytsearch1000:#shorts #handmadegifts",
    "ytsearch1000:#shorts #lifehacks",
    "ytsearch1000:#shorts #creative",
    "ytsearch1000:#shorts #origami",
    "ytsearch1000:#shorts #papercraft",
    "ytsearch1000:#shorts #diyprojects",
    "ytsearch1000:#shorts #woodworking",
    "ytsearch1000:#shorts #diyhacks",
    "ytsearch1000:#shorts #artandcraft",
    "ytsearch1000:#shorts #diydecor",
    "ytsearch1000:#shorts #recycling",
    "ytsearch1000:#shorts #upcycling",
    "ytsearch1000:#shorts #claycraft",
    "ytsearch1000:#shorts #knitting",
    "ytsearch1000:#shorts #crochet",
    "ytsearch1000:#shorts #sewing",
    "ytsearch1000:#shorts #embroidery",
    "ytsearch1000:#shorts #painting",
    "ytsearch1000:#shorts #drawing",
    "ytsearch1000:#shorts #cardmaking",
    "ytsearch1000:#shorts #resinart",
    "ytsearch1000:#shorts #jewelrymaking",
    "ytsearch1000:#shorts #miniature",
]


def load_pool() -> dict:
    if os.path.exists(POOL_FILE):
        with open(POOL_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_pool(pool: dict):
    os.makedirs(os.path.dirname(POOL_FILE), exist_ok=True)
    with open(POOL_FILE, "w", encoding="utf-8") as f:
        json.dump(pool, f, ensure_ascii=False, indent=2)
    _update_stats(pool)


def _update_stats(pool: dict):
    total   = len(pool)
    unused  = sum(1 for v in pool.values() if not v.get("used") and not v.get("failed"))
    used    = sum(1 for v in pool.values() if v.get("used"))
    failed  = sum(1 for v in pool.values() if v.get("failed"))
    by_tag  = {}
    for v in pool.values():
        tag = v.get("hashtag", "?")
        by_tag[tag] = by_tag.get(tag, 0) + 1

    stats = {
        "total": total, "unused": unused, "used": used, "failed": failed,
        "by_hashtag": by_tag,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    with open(STATS_FILE, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)
    print(f"\n📊 [VIDEO POOL THỐNG KÊ] Tổng: {total:,} video | Unused: {unused:,} | Used: {used} | Failed: {failed}")


def fetch_query(query: str) -> list[dict]:
    cmd = [
        "yt-dlp",
        "--flat-playlist",
        "--print", "%(id)s\t%(webpage_url)s\t%(title)s\t%(duration)s",
        "--no-warnings",
        "--quiet",
        query,
    ]
    entries = []
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        lines = [l.strip() for l in res.stdout.strip().splitlines() if "\t" in l]

        tag_name = "#diy" if "diy" in query else "#handmade"

        for line in lines:
            parts = line.split("\t", 3)
            if len(parts) < 2:
                continue
            vid_id  = parts[0].strip()
            vid_url = parts[1].strip()
            title   = parts[2].strip() if len(parts) > 2 else ""

            entries.append({
                "id":       vid_id,
                "url":      vid_url,
                "platform": "youtube_shorts",
                "hashtag":  tag_name,
                "title":    title[:100],
                "added_at": datetime.now().strftime("%Y-%m-%d"),
                "used":     False,
                "failed":   False,
            })
        print(f"  ✅ Query '{query[:35]}...' → Thu thập {len(entries)} video link")
    except Exception as e:
        print(f"  ❌ Query '{query[:35]}...' lỗi: {e}")
    return entries


def run_batch_video_crawl(max_threads: int = 10):
    pool = load_pool()
    initial_total = len(pool)
    new_added = 0

    print(f"🚀 [MASSIVE VIDEO CRAWLER] Quét 10,000 - 20,000 link video dọc qua {len(HASHTAG_QUERIES)} truy vấn...")
    print(f"⚡ Số luồng xử lý đồng thời: {max_threads} threads\n")
    start_time = time.time()

    with ThreadPoolExecutor(max_workers=max_threads) as executor:
        futures = {executor.submit(fetch_query, q): q for q in HASHTAG_QUERIES}

        for future in as_completed(futures):
            items = future.result()
            for item in items:
                vid_id = item["id"]
                if vid_id not in pool:
                    pool[vid_id] = item
                    new_added += 1

            save_pool(pool)

    elapsed = time.time() - start_time
    print(f"\n🎉 [HOÀN THÀNH MASSIVE VIDEO CRAWLER] Thu thập thêm {new_added:,} link video mới trong {elapsed:.1f}s. TỔNG KHO VIDEO POOL: {len(pool):,} video!")
    return new_added


if __name__ == "__main__":
    run_batch_video_crawl(max_threads=10)
