"""
story_scraper.py
================
Scrape nội dung câu chuyện từ các nguồn:
- URL bất kỳ (generic HTML scraper)
- Reddit (dùng PRAW API)
- Các trang Việt Nam phổ biến

Cách dùng:
    python story_scraper.py --url "https://..." --output story.json
    python story_scraper.py --reddit --subreddit tifu --limit 5
"""

import os
import re
import json
import argparse
import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
}
MAX_WORDS = 600   # ~60 giây TTS ở tốc độ 1.25x
MIN_WORDS = 80    # ít nhất 80 từ mới dùng


# ---------------------------------------------------------------------------
# GENERIC HTML SCRAPER
# ---------------------------------------------------------------------------
def scrape_url(url: str) -> dict | None:
    """
    Scrape tiêu đề + nội dung chính từ URL bất kỳ.
    Trả về dict {"title": str, "content": str, "url": str} hoặc None.
    """
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding
    except Exception as e:
        print(f"[Scraper] Lỗi GET {url}: {e}")
        return None

    soup = BeautifulSoup(resp.text, "lxml")

    # --- Lấy tiêu đề ---
    title = ""
    if soup.find("h1"):
        title = soup.find("h1").get_text(strip=True)
    elif soup.find("title"):
        title = soup.find("title").get_text(strip=True)

    # --- Thử các selector phổ biến để lấy nội dung chính ---
    content = ""
    selectors = [
        "article",
        '[class*="content"]',
        '[class*="article"]',
        '[class*="post-body"]',
        '[class*="entry"]',
        "main",
        ".story",
        "#content",
    ]
    for sel in selectors:
        el = soup.select_one(sel)
        if el:
            paragraphs = el.find_all("p")
            if paragraphs:
                content = " ".join(p.get_text(strip=True) for p in paragraphs)
                break

    # Fallback: lấy tất cả <p>
    if not content:
        paragraphs = soup.find_all("p")
        content = " ".join(p.get_text(strip=True) for p in paragraphs)

    content = _clean_text(content)

    if len(content.split()) < MIN_WORDS:
        print(f"[Scraper] Nội dung quá ngắn ({len(content.split())} từ), bỏ qua.")
        return None

    # Giới hạn độ dài
    content = _truncate_to_words(content, MAX_WORDS)

    return {"title": title, "content": content, "url": url}


# ---------------------------------------------------------------------------
# REDDIT SCRAPER  (không cần API key - dùng old.reddit JSON)
# ---------------------------------------------------------------------------
def scrape_reddit_post(post_url: str) -> dict | None:
    """
    Scrape 1 bài Reddit cụ thể qua old.reddit JSON API (không cần API key).
    """
    json_url = post_url.rstrip("/") + ".json"
    try:
        resp = requests.get(json_url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        post = data[0]["data"]["children"][0]["data"]
        title = post.get("title", "")
        content = post.get("selftext", "")
        if not content:
            content = title   # link post, không có text
        content = _clean_text(content)
        content = _truncate_to_words(content, MAX_WORDS)
        return {"title": title, "content": content, "url": post_url}
    except Exception as e:
        print(f"[Reddit] Lỗi scrape {post_url}: {e}")
        return None


def get_top_reddit_stories(subreddit: str = "tifu", limit: int = 5,
                           time_filter: str = "day") -> list[dict]:
    """
    Lấy top N bài từ subreddit qua Reddit JSON API (không cần OAuth).
    time_filter: hour | day | week | month | year | all
    """
    url = f"https://www.reddit.com/r/{subreddit}/top.json?limit={limit}&t={time_filter}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        posts = resp.json()["data"]["children"]
        results = []
        for p in posts:
            d = p["data"]
            if d.get("is_video") or not d.get("selftext"):
                continue
            content = _clean_text(d.get("selftext", ""))
            if len(content.split()) < MIN_WORDS:
                continue
            results.append({
                "title": d["title"],
                "content": _truncate_to_words(content, MAX_WORDS),
                "url": f"https://reddit.com{d['permalink']}",
                "score": d.get("score", 0),
            })
        return results
    except Exception as e:
        print(f"[Reddit] Lỗi get top: {e}")
        return []


# ---------------------------------------------------------------------------
# HELPER
# ---------------------------------------------------------------------------
def _clean_text(text: str) -> str:
    """Loại bỏ markdown, nhiều khoảng trắng, emoji thô..."""
    text = re.sub(r"\*+", "", text)          # bold/italic markdown
    text = re.sub(r"#{1,6}\s", "", text)     # headings markdown
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)   # links
    text = re.sub(r"https?://\S+", "", text) # URLs
    text = re.sub(r"#\w+\s*", "", text)      # hashtags
    # Emoji unicode range
    emoji_pattern = re.compile(
        "[\U0001F600-\U0001F64F\U0001F300-\U0001F5FF"
        "\U0001F680-\U0001F6FF\U0001F1E0-\U0001F1FF"
        "\U00002702-\U000027B0\U000024C2-\U0001F251]+",
        flags=re.UNICODE,
    )
    text = emoji_pattern.sub("", text)
    text = re.sub(r"\n+", " ", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def _truncate_to_words(text: str, max_words: int) -> str:
    words = text.split()
    if len(words) <= max_words:
        return text
    # Tìm dấu chấm gần nhất để câu không bị cắt nửa chừng
    truncated = " ".join(words[:max_words])
    last_dot = max(truncated.rfind("."), truncated.rfind("!"), truncated.rfind("?"))
    if last_dot > len(truncated) // 2:
        return truncated[: last_dot + 1]
    return truncated + "..."


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Story Scraper")
    parser.add_argument("--url", help="URL trang web cần scrape")
    parser.add_argument("--reddit", action="store_true", help="Lấy từ Reddit")
    parser.add_argument("--subreddit", default="tifu", help="Tên subreddit")
    parser.add_argument("--limit", type=int, default=3)
    parser.add_argument("--output", default="story.json")
    args = parser.parse_args()

    if args.reddit:
        stories = get_top_reddit_stories(args.subreddit, args.limit)
        if stories:
            with open(args.output, "w", encoding="utf-8") as f:
                json.dump(stories[0], f, ensure_ascii=False, indent=2)
            print(f"[OK] Đã lưu story vào {args.output}")
            print(f"Title: {stories[0]['title']}")
        else:
            print("[FAIL] Không lấy được story nào.")
    elif args.url:
        story = scrape_url(args.url)
        if story:
            with open(args.output, "w", encoding="utf-8") as f:
                json.dump(story, f, ensure_ascii=False, indent=2)
            print(f"[OK] Đã lưu story vào {args.output}")
        else:
            print("[FAIL] Không lấy được nội dung.")
    else:
        print("Dùng --url <URL> hoặc --reddit --subreddit <sub>")
