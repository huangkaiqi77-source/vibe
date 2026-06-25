"""
微博评论爬虫 —— 面向女性运动员数字形象研究
=============================================
基于 m.weibo.cn 移动端 API，按运动员关键词搜索帖子并采集评论。

使用前准备：
  1. 浏览器打开 https://m.weibo.cn 并登录
  2. F12 → Application → Cookies → 复制 SUB 的值
  3. 填入下方 COOKIE_CONFIG['SUB']

运行：
  python weibo_crawler.py

特性：
  - 断点续传：中断后重新运行会跳过已采集的帖子
  - 自动限速：请求间随机延迟，降低被封概率
  - 数据导出：CSV（Excel 可直接打开）+ JSON 双格式
"""

import requests
import time
import random
import json
import csv
import os
import re
import sys
from datetime import datetime
from urllib.parse import quote

# ---- 修复 Windows 终端 GBK 编码无法输出 Emoji 的问题 ----
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ========================== 配置区域 ==========================

# 微博 Cookie —— 至少需要填写 SUB 字段
COOKIE_CONFIG = {
    "SUB": "_2A25HPtvIDeRhGeNM71QQ8y_LzDiIHXVkMlEArDV6PUJbktANLVD2kW1NTg_adHyWYDVxa2cviT4666_gFFCBedY1",
    "SUBP": "",
    "WEIBOCN_FROM": "1110006030",
}

# 目标运动员关键词列表
ATHLETE_KEYWORDS = [
    # 女性运动员
    "孙颖莎",
    "全红婵",
    "谷爱凌",
    # 男性运动员（对照组）
    "马龙",
    "樊振东",
    "苏炳添",
]

# 每个关键词搜索的最大页数（每页约 10-20 条帖子）
MAX_SEARCH_PAGES = 10

# 每个帖子采集的最大评论页数（微博评论每页约 20 条）
MAX_COMMENT_PAGES = 15

# 请求间隔（秒），会在 [min, max] 区间内随机
REQUEST_DELAY_MIN = 2.0
REQUEST_DELAY_MAX = 5.0

# 遇到限流时的冷却时间（秒）
RATE_LIMIT_COOLDOWN = 120

# 最大重试次数
MAX_RETRIES = 3

# 输出目录
OUTPUT_DIR = "data"

# ========================== 工具函数 ==========================


def build_cookie_string():
    """将 COOKIE_CONFIG 拼接成 Cookie 字符串"""
    parts = []
    for key, value in COOKIE_CONFIG.items():
        if value:
            parts.append(f"{key}={value}")
    return "; ".join(parts)


def build_headers():
    """构造请求头"""
    return {
        "User-Agent": (
            "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) "
            "AppleWebKit/605.1.15 (KHTML, like Gecko) "
            "Version/16.0 Mobile/15E148 Safari/604.1"
        ),
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Referer": "https://m.weibo.cn/",
        "Cookie": build_cookie_string(),
        "X-Requested-With": "XMLHttpRequest",
    }


def random_delay():
    """随机休眠，降低请求频率"""
    delay = random.uniform(REQUEST_DELAY_MIN, REQUEST_DELAY_MAX)
    time.sleep(delay)


def safe_request(url, params=None, retries=MAX_RETRIES):
    """带重试和限流检测的请求封装"""
    for attempt in range(retries):
        try:
            resp = requests.get(
                url,
                params=params,
                headers=build_headers(),
                timeout=15,
            )

            # 被限流
            if resp.status_code == 418 or resp.status_code == 429:
                print(f"  ⚠️ 触发限流（{resp.status_code}），冷却 {RATE_LIMIT_COOLDOWN}s...")
                time.sleep(RATE_LIMIT_COOLDOWN)
                continue

            # 需要登录
            if resp.status_code == 403 or "请先登录" in resp.text:
                print("  ❌ Cookie 无效或已过期，请重新获取 SUB 值")
                return None

            if resp.status_code != 200:
                print(f"  ⚠️ HTTP {resp.status_code}，第 {attempt + 1} 次重试...")
                time.sleep(5 * (attempt + 1))
                continue

            try:
                return resp.json()
            except json.JSONDecodeError:
                print(f"  ⚠️ 响应非 JSON，第 {attempt + 1} 次重试...")
                time.sleep(3)
                continue

        except requests.RequestException as e:
            print(f"  ⚠️ 网络错误: {e}，第 {attempt + 1} 次重试...")
            time.sleep(5 * (attempt + 1))

    print(f"  ❌ 请求最终失败: {url}")
    return None


def clean_text(text):
    """清洗文本：去除 HTML 标签、多余空白、特殊字符"""
    if not text:
        return ""
    # 去除 HTML 标签
    text = re.sub(r"<[^>]+>", "", text)
    # 去除 \\n、\\t 等转义
    text = text.replace("\\n", " ").replace("\\t", " ").replace("\\r", " ")
    # 合并多个空白
    text = re.sub(r"\s+", " ", text)
    # 去除首尾空白
    return text.strip()


def ensure_output_dir():
    """确保输出目录存在"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)


def timestamp_to_date(ts):
    """微博时间戳或日期字符串 → 统一格式 YYYY-MM-DD HH:MM:SS"""
    if not ts:
        return ""
    # 已经是字符串格式（如 "Thu Jan 15 12:30:00 +0800 2025" 或 "2025-01-15"）
    if isinstance(ts, str):
        return clean_text(ts)
    # Unix 时间戳（整数或浮点）
    try:
        ts_int = int(ts)
        return datetime.fromtimestamp(ts_int).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, OSError, TypeError):
        return str(ts)


# ========================== 搜索帖子 ==========================


def search_posts(keyword, max_pages=MAX_SEARCH_PAGES):
    """
    按关键词搜索微博帖子。
    返回: [{post_id, user_name, content, likes, comments_count, reposts, created_at, keyword}, ...]
    """
    posts = []
    seen_in_keyword = set()  # 当前关键词内去重
    page = 1
    # 初始 containerid：用 quote 编码关键字
    next_containerid = f"100103type%3D1%26q%3D{quote(keyword)}"

    while page <= max_pages:
        url = f"https://m.weibo.cn/api/container/getIndex"
        params = {
            "containerid": next_containerid,
            "page_type": "searchall",
        }

        print(f"  搜索「{keyword}」第 {page}/{max_pages} 页...")
        data = safe_request(url, params=params)
        random_delay()

        if data is None:
            print(f"    搜索中断于第 {page} 页")
            break

        if data.get("ok") != 1:
            print(f"    API 返回异常: {data.get('msg', '未知错误')}")
            break

        cards = data.get("data", {}).get("cards", [])
        if not cards:
            print(f"    第 {page} 页无结果，搜索结束")
            break

        page_new = 0
        for card in cards:
            if card.get("card_type") != 9:
                continue

            mblog = card.get("mblog")
            if not mblog:
                continue

            post_id = mblog.get("id", "")
            if post_id in seen_in_keyword:
                continue
            seen_in_keyword.add(post_id)

            user = mblog.get("user", {})
            user_name = user.get("screen_name", "未知") if user else "未知"

            # 过滤：至少要有一定评论量的帖子才值得采集
            comments_count = mblog.get("comments_count", 0)
            if comments_count < 3:
                continue

            content = clean_text(mblog.get("text", ""))
            created_at = timestamp_to_date(mblog.get("created_at", 0))

            posts.append({
                "post_id": post_id,
                "user_name": user_name,
                "content": content[:200],
                "likes": mblog.get("attitudes_count", 0),
                "comments_count": comments_count,
                "reposts": mblog.get("reposts_count", 0),
                "created_at": created_at,
                "keyword": keyword,
            })
            page_new += 1

        print(f"    本页采集 {page_new} 条新帖子")

        # 获取下一页的 containerid（API 返回的真实翻页 token）
        cardlist_info = data.get("data", {}).get("cardlistInfo", {})
        new_cid = cardlist_info.get("containerid", "")
        if new_cid and new_cid != next_containerid:
            next_containerid = new_cid
            page += 1
        elif page_new == 0:
            # 既无新 containerid 也无新帖子，结束
            print(f"    无更多页面")
            break
        else:
            page += 1

    print(f"  「{keyword}」共搜索到 {len(posts)} 条帖子")
    return posts


# ========================== 采集评论 ==========================


def fetch_comments(post_id, max_pages=MAX_COMMENT_PAGES):
    """
    获取单条帖子的评论。
    优先使用 hotflow（热门评论），数据更有代表性。
    返回: [{comment_id, user_name, content, likes, created_at, post_id}, ...]
    """
    comments = []
    max_id = 0

    for page in range(max_pages):
        url = "https://m.weibo.cn/comments/hotflow"
        params = {
            "id": post_id,
            "mid": post_id,
            "max_id": max_id,
            "max_id_type": 0,
        }

        data = safe_request(url, params=params)
        random_delay()

        if data is None:
            break

        if data.get("ok") != 1:
            # hotflow 失败，尝试普通评论接口
            alt_data = _fetch_comments_alt(post_id, page + 1)
            if alt_data is None:
                break
            data = alt_data

        # 提取评论数据
        comment_list = data.get("data", {}).get("data", [])
        if not comment_list:
            break

        for item in comment_list:
            user = item.get("user", {})
            comment = {
                "comment_id": item.get("id", ""),
                "user_name": user.get("screen_name", "匿名") if user else "匿名",
                "content": clean_text(item.get("text", "")),
                "likes": item.get("like_count", 0),
                "created_at": timestamp_to_date(item.get("created_at", 0)),
                "post_id": post_id,
            }
            # 过滤空评论
            if comment["content"]:
                comments.append(comment)

        # 更新翻页游标
        new_max_id = data.get("data", {}).get("max_id", 0)
        if new_max_id == 0 or new_max_id == max_id:
            break
        max_id = new_max_id

    return comments


def _fetch_comments_alt(post_id, page):
    """备用评论接口（按时间排序）"""
    url = "https://m.weibo.cn/api/comments/show"
    params = {"id": post_id, "page": page}
    return safe_request(url, params=params)


# ========================== 断点续传 ==========================


def save_checkpoint(data, filename):
    """保存中间数据为 JSON（断点续传用）"""
    ensure_output_dir()
    filepath = os.path.join(OUTPUT_DIR, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_checkpoint(filename):
    """读取已有的中间数据"""
    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        return None
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


# ========================== 导出 ==========================


def export_csv(comments, posts, prefix="weibo"):
    """导出为 CSV 文件"""
    ensure_output_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 导出帖子
    posts_file = os.path.join(OUTPUT_DIR, f"{prefix}_posts_{timestamp}.csv")
    with open(posts_file, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "post_id", "keyword", "user_name", "content",
            "likes", "comments_count", "reposts", "created_at",
        ])
        writer.writeheader()
        for p in posts:
            writer.writerow(p)

    # 导出评论
    comments_file = os.path.join(OUTPUT_DIR, f"{prefix}_comments_{timestamp}.csv")
    with open(comments_file, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "comment_id", "post_id", "user_name", "content",
            "likes", "created_at",
        ])
        writer.writeheader()
        for c in comments:
            writer.writerow(c)

    print(f"\n📁 帖子数据: {posts_file}  ({len(posts)} 条)")
    print(f"📁 评论数据: {comments_file}  ({len(comments)} 条)")


def export_json(comments, posts, prefix="weibo"):
    """同时导出 JSON 格式（方便程序处理）"""
    ensure_output_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    all_data = {
        "metadata": {
            "crawled_at": timestamp,
            "total_comments": len(comments),
            "total_posts": len(posts),
            "keywords": ATHLETE_KEYWORDS,
        },
        "posts": posts,
        "comments": comments,
    }

    filepath = os.path.join(OUTPUT_DIR, f"{prefix}_full_{timestamp}.json")
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)

    print(f"📁 完整数据: {filepath}")


# ========================== 主流程 ==========================


def main():
    print("=" * 60)
    print("  微博评论爬虫 — 女性运动员数字形象研究")
    print("=" * 60)
    print()

    # 检查 Cookie
    if not COOKIE_CONFIG["SUB"]:
        print("❌ 请先配置 Cookie！")
        print()
        print("   操作步骤：")
        print("   1. 浏览器打开 https://m.weibo.cn 并登录")
        print("   2. F12 → Application → Cookies → m.weibo.cn")
        print("   3. 找到 SUB 这一行，复制它的 Value")
        print("   4. 粘贴到脚本顶部 COOKIE_CONFIG['SUB'] = '你的SUB值'")
        print()
        sys.exit(1)

    print(f"🎯 目标运动员: {', '.join(ATHLETE_KEYWORDS)}")
    print(f"📊 搜索页数上限: {MAX_SEARCH_PAGES} 页/关键词")
    print(f"💬 评论页数上限: {MAX_COMMENT_PAGES} 页/帖子")
    print()

    # 加载已有的采集进度
    all_posts = load_checkpoint("checkpoint_posts.json") or []
    all_comments = load_checkpoint("checkpoint_comments.json") or []
    crawled_post_ids = {c["post_id"] for c in all_comments}

    if all_posts:
        print(f"📂 断点续传：已有 {len(all_posts)} 条帖子，{len(all_comments)} 条评论")
        print()

    # ---- 阶段 1：搜索帖子 ----
    print("=" * 40)
    print("  阶段 1/2：搜索帖子")
    print("=" * 40)

    known_post_ids = {p["post_id"] for p in all_posts}

    for keyword in ATHLETE_KEYWORDS:
        new_posts = search_posts(keyword, MAX_SEARCH_PAGES)
        for post in new_posts:
            if post["post_id"] not in known_post_ids:
                known_post_ids.add(post["post_id"])
                all_posts.append(post)
        # 每个关键词搜完立刻存盘，防止中断丢失
        save_checkpoint(all_posts, "checkpoint_posts.json")
    print(f"\n📌 帖子搜索完成，共 {len(all_posts)} 条")
    print()

    # ---- 阶段 2：采集评论 ----
    print("=" * 40)
    print("  阶段 2/2：采集评论")
    print("=" * 40)

    # 按评论数降序排列，优先爬热帖
    posts_to_crawl = sorted(
        all_posts,
        key=lambda p: p["comments_count"],
        reverse=True,
    )

    total = len(posts_to_crawl)
    for idx, post in enumerate(posts_to_crawl):
        pid = post["post_id"]
        keyword = post["keyword"]

        if pid in crawled_post_ids:
            continue

        print(f"\n[{idx + 1}/{total}] 帖子 {pid[:12]}... "
              f"|「{keyword}」| 已有 {post['comments_count']} 条评论")

        comments = fetch_comments(pid, MAX_COMMENT_PAGES)

        if comments:
            all_comments.extend(comments)
            crawled_post_ids.add(pid)
            print(f"  ✅ 采集到 {len(comments)} 条评论（累计 {len(all_comments)} 条）")

            # 每 5 个帖子保存一次中间结果
            if (idx + 1) % 5 == 0:
                save_checkpoint(all_comments, "checkpoint_comments.json")
        else:
            print(f"  ⏭️ 无评论或获取失败，跳过")

        # 每 50 个帖子停顿更久一点
        if (idx + 1) % 50 == 0:
            print("\n  😴 长休息 30 秒...")
            time.sleep(30)

    # ---- 最终保存 ----
    save_checkpoint(all_comments, "checkpoint_comments.json")

    print()
    print("=" * 60)
    print("  采集完成！")
    print("=" * 60)
    print(f"  帖子总数: {len(all_posts)}")
    print(f"  评论总数: {len(all_comments)}")

    # 按关键词统计
    for keyword in ATHLETE_KEYWORDS:
        kw_comments = [c for c in all_comments
                       for p in all_posts
                       if p["post_id"] == c["post_id"] and p["keyword"] == keyword]
        print(f"    「{keyword}」: {len(kw_comments)} 条评论")

    print()

    # ---- 导出 ----
    export_csv(all_comments, all_posts)
    export_json(all_comments, all_posts)

    print("\n✨ 完成！数据已保存到 ./data/ 目录")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️ 手动中断。已采集的数据已通过断点保存，重新运行即可续传。")
        sys.exit(0)
