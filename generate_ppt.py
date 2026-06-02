import os
import requests
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from openai import OpenAI
from datetime import datetime
import feedparser

# 读取配置
DEEPSEEK_API_KEY = os.environ['DEEPSEEK_API_KEY']
WECOM_WEBHOOK = os.environ['WECOM_WEBHOOK']
NEWS_SOURCE = os.environ.get('NEWS_SOURCE', '科技,财经')

# DeepSeek客户端
client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

def fetch_news(url=None):
    """从指定 US News 页面抓取新闻标题与链接"""
    if not url:
        url = "https://www.usnews.com/topics/subjects/wall-street"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
    except Exception as e:
        print(f"请求网页失败: {e}")
        return []

    from bs4 import BeautifulSoup
    soup = BeautifulSoup(resp.text, 'lxml')

    # US News 的新闻列表通常在 <div> 或者 <article> 中，选择器可能随改版变化
    # 当前常见结构：<a class="Anchor-sc-..."> 包含标题
    articles = []
    # 尝试找到所有带标题的链接
    for a_tag in soup.select('a[href]'):
        title = a_tag.get_text(strip=True)
        href = a_tag.get('href')
        # 过滤掉太短、非新闻标题的链接
        if title and len(title) > 25 and '/articles/' in href:
            full_url = href if href.startswith('http') else f'https://www.usnews.com{href}'
            articles.append(f"{title} (来源: {full_url})")
        if len(articles) >= 10:
            break

    if not articles:
        print("未找到新闻，可能网站结构已变化，请调整选择器。")
    return articles

def summarize_news(articles):
    news_text = "\n".join(articles)
    prompt = f"""你是一个新闻编辑，请将以下新闻概括为一份今日要闻简报，要求：
    - 口语化，适合朗读
    - 分成3~5个要点，每个要点一句话概括
    - 开头加上日期和一句话导语
    新闻内容：
    {news_text}
    输出格式直接是简报文本，不要多余的话。"""
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
    )
    return response.choices[0].message.content

def create_ppt(brief_text, date_str):
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    
    title.text = f"{date_str} 每日新闻简报"
    content.text = brief_text
    
    for paragraph in content.text_frame.paragraphs:
        paragraph.font.size = Pt(18)
        paragraph.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    ppt_filename = "daily_news.pptx"
    prs.save(ppt_filename)
    return ppt_filename

def send_to_wecom(file_path, webhook):
    key = webhook.split('key=')[-1]
    upload_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={key}&type=file"
    with open(file_path, 'rb') as f:
        files = {'media': f}
        r = requests.post(upload_url, files=files)
    media_id = r.json()['media_id']
    
    data = {
        "msgtype": "file",
        "file": {"media_id": media_id}
    }
    requests.post(webhook, json=data)

if __name__ == '__main__':
    keywords = NEWS_SOURCE
    articles = fetch_news(keywords)
    if not articles:
        print("没有抓到新闻，退出")
        exit(1)
    brief = summarize_news(articles)
    today = datetime.now().strftime("%Y年%m月%d日")
    ppt_file = create_ppt(brief, today)
    send_to_wecom(ppt_file, WECOM_WEBHOOK)
    print("推送完成")
