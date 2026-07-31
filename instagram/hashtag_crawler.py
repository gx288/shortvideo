"""
instagram/hashtag_crawler.py
============================
Thu thập link video từ hashtag TikTok / Instagram bằng yt-dlp --flat-playlist.
Lưu vào link_pool.json (chỉ URL, không download video).

Mục tiêu: 10,000+ link → mỗi lần tạo Shorts thì pick ngẫu nhiên 1-2 link rồi download.

Nguồn:
    TikTok:    https://www.tiktok.com/tag/<hashtag>   ✅ Hoạt động tốt
    Instagram: https://www.instagram.com/explore/tags/<hashtag>/  ⚠️ Cần cookie

Cách dùng:
    python instagram/hashtag_crawler.py --hashtag diy funny lifehack --limit 500
    python instagram/hashtag_crawler.py --hashtag diy --platform tiktok --limit 2000
"""

import os
import sys
import json
import time
import random
import argparse
import subprocess
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

POOL_FILE   = os.path.join("instagram", "link_pool.json")
STATS_FILE  = os.path.join("instagram", "pool_stats.json")

# ─────────────────────────────────────────────────────────────────────────────
# POOL I/O
# ─────────────────────────────────────────────────────────────────────────────

def load_pool() -> dict:
    """
    Pool structure:
    {
      "<video_id>": {
        "url":       "https://...",
        "platform":  "tiktok" | "instagram",
        "hashtag":   "#diy",
        "title":     "...",
        "added_at":  "2026-07-30",
        "used":      false,
        "failed":    false
      }, ...
    }
    """
    if os.path.exists(POOL_FILE):
        with open(POOL_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_pool(pool: dict):
    os.makedirs(os.path.dirname(POOL_FILE), exist_ok=True)
    with open(POOL_FILE, "w", encoding="utf-8") as f:
        json.dump(pool, f, ensure_ascii=False, separators=(",", ":"))
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
    print(f"\n📊 Pool: {total} total | {unused} unused | {used} used | {failed} failed")
    for tag, cnt in sorted(by_tag.items(), key=lambda x: -x[1]):
        print(f"   {tag}: {cnt}")


# ─────────────────────────────────────────────────────────────────────────────
# TIKTOK HASHTAG CRAWLER
# ─────────────────────────────────────────────────────────────────────────────

def crawl_tiktok_hashtag(hashtag: str, limit: int = 500) -> list[dict]:
    """
    Dùng yt-dlp --flat-playlist để lấy danh sách URL từ hashtag TikTok.
    Không download video, chỉ lấy metadata.
    """
    tag = hashtag.lstrip("#")
    url = f"https://www.tiktok.com/tag/{tag}"
    print(f"\n[TikTok] Crawling #{tag} (limit={limit})...")

    cmd = [
        "yt-dlp",
        "--flat-playlist",
        "--playlist-end", str(limit),
        "--print", "%(id)s\t%(webpage_url)s\t%(title)s\t%(duration)s\t%(view_count)s",
        "--no-warnings",
        "--quiet",
        "--extractor-args", "tiktok:api_hostname=api16-normal-c-useast1a.tiktokv.com",
        url,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=300)
        lines = [l.strip() for l in result.stdout.strip().splitlines() if "\t" in l]

        entries = []
        for line in lines:
            parts = line.split("\t", 4)
            if len(parts) < 2:
                continue
            vid_id   = parts[0].strip()
            vid_url  = parts[1].strip()
            title    = parts[2].strip() if len(parts) > 2 else ""
            duration = parts[3].strip() if len(parts) > 3 else ""
            views_str = parts[4].strip() if len(parts) > 4 else "0"

            # YÊU CẦU: Thời lượng tầm 30s (15s -> 45s) và view > 100k
            try:
                dur = float(duration)
                if dur < 15 or dur > 45:
                    continue
                
                # Check view_count
                views = int(float(views_str)) if views_str and views_str != "NA" else 0
                if views < 100000:
                    continue
            except (ValueError, TypeError):
                continue

            entries.append({
                "id":       vid_id,
                "url":      vid_url,
                "platform": "tiktok",
                "hashtag":  f"#{tag}",
                "title":    title[:100],
                "added_at": datetime.now().strftime("%Y-%m-%d"),
                "used":     False,
                "failed":   False,
            })

        print(f"[TikTok] ✅ Thu thập được {len(entries)} link hợp lệ từ #{tag}")
        if not entries:
            print(f"[TikTok] ⚠️ TikTok tag bị khoá/lỗi, chuyển sang YouTube Shorts cho #{tag}...")
            return crawl_youtube_shorts(tag, limit)
        return entries

    except Exception as e:
        print(f"[TikTok] ❌ Lỗi: {e}, fallback YouTube Shorts...")
        return crawl_youtube_shorts(tag, limit)


# ─────────────────────────────────────────────────────────────────────────────
# YOUTUBE SHORTS CRAWLER (Fallback cực kỳ ổn định, không lo bị khoá)
# ─────────────────────────────────────────────────────────────────────────────

def crawl_youtube_shorts(hashtag: str, limit: int = 500) -> list[dict]:
    """
    Dùng yt-dlp search `#shorts #{hashtag}` để lấy video dọc chất lượng cao.
    """
    tag = hashtag.lstrip("#")
    search_query = f"ytsearch{limit}:#shorts #{tag}"
    print(f"[YouTube Shorts] Crawling #{tag} (limit={limit})...")

    cmd = [
        "yt-dlp",
        "--flat-playlist",
        "--playlist-end", str(limit),
        "--print", "%(id)s\t%(webpage_url)s\t%(title)s\t%(duration)s\t%(view_count)s",
        "--no-warnings",
        "--quiet",
        search_query,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=300)
        lines = [l.strip() for l in result.stdout.strip().splitlines() if "\t" in l]

        entries = []
        for line in lines:
            parts = line.split("\t", 4)
            if len(parts) < 2:
                continue
            vid_id   = parts[0].strip()
            vid_url  = parts[1].strip()
            title    = parts[2].strip() if len(parts) > 2 else ""
            duration = parts[3].strip() if len(parts) > 3 else ""
            views_str = parts[4].strip() if len(parts) > 4 else "0"

            try:
                dur = float(duration)
                if dur < 15 or dur > 45:
                    continue
                
                # Check view_count
                views = int(float(views_str)) if views_str and views_str != "NA" else 0
                if views < 100000:
                    continue
            except (ValueError, TypeError):
                continue

            entries.append({
                "id":       vid_id,
                "url":      vid_url,
                "platform": "youtube_shorts",
                "hashtag":  f"#{tag}",
                "title":    title[:100],
                "added_at": datetime.now().strftime("%Y-%m-%d"),
                "used":     False,
                "failed":   False,
            })

        print(f"[YouTube Shorts] ✅ Thu thập được {len(entries)} link từ #{tag}")
        return entries

    except Exception as e:
        print(f"[YouTube Shorts] ❌ Lỗi: {e}")
        return []



# ─────────────────────────────────────────────────────────────────────────────
# INSTAGRAM HASHTAG CRAWLER (cần cookie)
# ─────────────────────────────────────────────────────────────────────────────

def crawl_instagram_hashtag(hashtag: str, limit: int = 200,
                             cookie_file: str = None) -> list[dict]:
    """
    Crawl hashtag Instagram. Cần cookie file để Instagram không block.
    cookie_file: path đến file cookie Netscape format (export từ browser).
    """
    tag = hashtag.lstrip("#")
    url = f"https://www.instagram.com/explore/tags/{tag}/"
    print(f"\n[Instagram] Crawling #{tag} (limit={limit})...")

    cmd = [
        "yt-dlp",
        "--flat-playlist",
        "--playlist-end", str(limit),
        "--print", "%(id)s\t%(webpage_url)s\t%(title)s",
        "--no-warnings",
        "--quiet",
    ]
    if cookie_file and os.path.exists(cookie_file):
        cmd += ["--cookies", cookie_file]
    cmd.append(url)

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=180)
        lines = [l.strip() for l in result.stdout.strip().splitlines() if "\t" in l]

        if not lines and result.stderr:
            print(f"[Instagram] ⚠️ {result.stderr[:300]}")
            print("[Instagram] 💡 Instagram thường yêu cầu cookie để crawl hashtag.")
            return []

        entries = []
        for line in lines:
            parts = line.split("\t", 2)
            if len(parts) < 2:
                continue
            entries.append({
                "id":       parts[0].strip(),
                "url":      parts[1].strip(),
                "platform": "instagram",
                "hashtag":  f"#{tag}",
                "title":    parts[2].strip()[:100] if len(parts) > 2 else "",
                "added_at": datetime.now().strftime("%Y-%m-%d"),
                "used":     False,
                "failed":   False,
            })

        print(f"[Instagram] ✅ Thu thập được {len(entries)} link từ #{tag}")
        return entries

    except subprocess.TimeoutExpired:
        print(f"[Instagram] ⏱️ Timeout khi crawl #{tag}")
        return []
    except Exception as e:
        print(f"[Instagram] ❌ Lỗi: {e}")
        return []


# ─────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def crawl_hashtags(hashtags: list[str], platform: str = "tiktok",
                   limit_per_tag: int = 500, cookie_file: str = None) -> int:
    """
    Crawl nhiều hashtag, gộp vào pool, bỏ qua duplicate.
    Trả về số link MỚI thêm vào pool.
    """
    pool = load_pool()
    existing_ids = set(pool.keys())
    new_count = 0

    for tag in hashtags:
        if platform == "tiktok":
            entries = crawl_tiktok_hashtag(tag, limit_per_tag)
        else:
            entries = crawl_instagram_hashtag(tag, limit_per_tag, cookie_file)

        for entry in entries:
            vid_id = entry["id"]
            if vid_id and vid_id not in existing_ids:
                pool[vid_id] = entry
                existing_ids.add(vid_id)
                new_count += 1

        # Nghỉ giữa các hashtag để không bị rate limit
        if len(hashtags) > 1:
            sleep_time = random.uniform(2, 5)
            time.sleep(sleep_time)

    save_pool(pool)
    print(f"\n✅ Thêm {new_count} link mới vào pool. Tổng: {len(pool)}")
    return new_count


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Hashtag Link Pool Crawler")
    parser.add_argument("--hashtag", nargs="+", required=True,
                        help="Một hoặc nhiều hashtag (có hoặc không có #)")
    parser.add_argument("--platform", choices=["tiktok", "instagram"], default="tiktok",
                        help="Nền tảng crawl (mặc định: tiktok)")
    parser.add_argument("--limit", type=int, default=500,
                        help="Số link tối đa mỗi hashtag (mặc định: 500)")
    parser.add_argument("--cookies", help="Path file cookie (cần cho Instagram)")
    parser.add_argument("--stats", action="store_true", help="Chỉ xem thống kê pool")
    args = parser.parse_args()

    if args.stats:
        pool = load_pool()
        _update_stats(pool)
    else:
        crawl_hashtags(
            hashtags=args.hashtag,
            platform=args.platform,
            limit_per_tag=args.limit,
            cookie_file=args.cookies,
        )
