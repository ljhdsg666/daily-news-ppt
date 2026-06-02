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

def fetch_news(keywords):
    articles = []
    for kw in keywords.split(','):
        url = f"https://news.google.com/rss/search?q={kw.strip()}&hl=zh-CN&gl=CN&ceid=CN:zh-Hans"
        feed = feedparser.parse(url)
        for entry in feed.entries[:5]:
            articles.append(entry.title + "：" + entry.summary[:200])
    return articles[:10]

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
