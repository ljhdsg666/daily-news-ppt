import os
import io
import json
import requests
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from openai import OpenAI
from datetime import datetime
import feedparser
from PIL import Image, ImageDraw

# ---------- 配置 ----------
DEEPSEEK_API_KEY = os.environ['DEEPSEEK_API_KEY']
WECOM_WEBHOOK = os.environ['WECOM_WEBHOOK']

client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

PRIMARY_COLOR = RGBColor(0x0D, 0x2C, 0x54)
ACCENT_COLOR = RGBColor(0xD4, 0xA8, 0x3C)
BG_LIGHT = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(255, 255, 255)

# ---------- 1. 抓取新闻（保留原文链接） ----------
def fetch_news():
    query = "wall street site:usnews.com"
    url = f"https://news.google.com/rss/search?q={requests.utils.quote(query)}&hl=en-US&gl=US&ceid=US:en"
    feed = feedparser.parse(url)
    articles = []
    for entry in feed.entries[:8]:
        # Google News RSS 的 link 字段是重定向后的地址，可以解析出真实链接
        # 但多数情况下这个链接就是原文
        articles.append({
            "title": entry.title,
            "link": entry.link
        })
    return articles

# ---------- 2. AI 生成结构化PPT方案（含详细内容） ----------
def generate_ppt_blueprint(articles, date_str):
    news_list = "\n".join([f"{i+1}. {a['title']}" for i, a in enumerate(articles)])
    prompt = f"""你是一个顶级咨询公司的演示文稿设计师。请根据以下今日华尔街新闻，设计一份专业PPT的完整方案。
要求：
- 输出严格JSON格式，不要有任何额外文字。
- 包含封面页、目录页、3~5条重点新闻的详细页、一个总结页。
- 每条新闻页必须包含：标题、4个要点（每个要点20~30字，尽量包含数据或影响分析）、图片搜索关键词（英文）。
- 不要输出链接，也不要编造链接。
- 总结页用3个精炼的句子概括今日重点。

JSON结构：
{{
  "slides": [
    {{"type": "cover", "title": "每日华尔街简报", "subtitle": "{date_str}"}},
    {{"type": "toc", "items": ["新闻标题1", "新闻标题2", ...]}},
    {{"type": "content", "title": "新闻标题", "points": ["要点1", "要点2", "要点3", "要点4"], "image_keyword": "finance chart"}},
    ...,
    {{"type": "summary", "text": "总结内容..."}}
  ]
}}

今日新闻：
{news_list}

请直接输出JSON。"""
    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.5,
            )
            raw = response.choices[0].message.content.strip()
            if raw.startswith("```json"):
                raw = raw[7:]
            if raw.endswith("```"):
                raw = raw[:-3]
            blueprint = json.loads(raw)
            slides = blueprint.get("slides", [])

            # 将真实链接按顺序附加到 content 页
            link_idx = 0
            for s in slides:
                if s.get("type") == "content" and link_idx < len(articles):
                    s["link"] = articles[link_idx]["link"]
                    link_idx += 1
            return slides
        except Exception as e:
            print(f"JSON解析失败，重试... {e}")
            continue
    print("AI生成PPT方案失败，使用默认结构")
    return [{"type": "cover", "title": "每日华尔街简报", "subtitle": date_str}]

# ---------- 3. 图片获取 ----------
def get_image(keyword):
    try:
        url = f"https://picsum.photos/800/600?random&{requests.utils.quote(keyword)}"
        resp = requests.get(url, timeout=8)
        if resp.status_code == 200 and 'image' in resp.headers.get('Content-Type',''):
            img = Image.open(io.BytesIO(resp.content))
            img_io = io.BytesIO()
            img.save(img_io, format='PNG')
            img_io.seek(0)
            return img_io
    except Exception as e:
        print(f"图片下载失败: {e}")
    # 质感渐变占位图
    img = Image.new('RGB', (800,600), (13,44,84))
    draw = ImageDraw.Draw(img)
    for i in range(600):
        color = (13 + i//15, 44 + i//20, 84 + i//25)
        draw.line([(0,i), (800,i)], fill=color)
    img_io = io.BytesIO()
    img.save(img_io, format='PNG')
    img_io.seek(0)
    return img_io

# ---------- 4. 生成专业PPT（含可点击链接）----------
def create_ppt_from_blueprint(slides_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def add_bg(slide, color=WHITE):
        bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()

    for slide_data in slides_data:
        stype = slide_data.get("type")

        # --- 封面页 ---
        if stype == "cover":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            img = get_image("wall street skyscraper")
            slide.shapes.add_picture(img, 0, 0, prs.slide_width, prs.slide_height)
            shade = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
            shade.fill.solid()
            shade.fill.fore_color.rgb = RGBColor(0,0,0)
            shade.fill.fore_color.brightness = -0.5
            shade.line.fill.background()
            txBox = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10), Inches(1.5))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = slide_data.get("title", "每日简报")
            p.font.size = Pt(54)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER
            txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(4.0), Inches(10), Inches(0.8))
            tf2 = txBox2.text_frame
            p2 = tf2.paragraphs[0]
            p2.text = slide_data.get("subtitle", "")
            p2.font.size = Pt(24)
            p2.font.color.rgb = ACCENT_COLOR
            p2.alignment = PP_ALIGN.CENTER

        # --- 目录页 ---
        elif stype == "toc":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(slide, WHITE)
            bar = slide.shapes.add_shape(1, 0, 0, Inches(0.8), prs.slide_height)
            bar.fill.solid()
            bar.fill.fore_color.rgb = PRIMARY_COLOR
            bar.line.fill.background()
            txBox = slide.shapes.add_textbox(Inches(1.5), Inches(0.8), Inches(10), Inches(0.8))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = "今日概览"
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_COLOR
            items = slide_data.get("items", [])
            y_start = 2.2
            for i, item in enumerate(items[:5]):
                txBox = slide.shapes.add_textbox(Inches(2), Inches(y_start + i*0.9), Inches(9), Inches(0.7))
                tf = txBox.text_frame
                p = tf.paragraphs[0]
                p.text = f"{i+1}.  {item}"
                p.font.size = Pt(22)
                p.font.color.rgb = RGBColor(50,50,50)

        # --- 内容页（详细要点 + 原文链接）---
        elif stype == "content":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(slide, BG_LIGHT)
            # 顶部标题栏
            header = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(1.2))
            header.fill.solid()
            header.fill.fore_color.rgb = PRIMARY_COLOR
            header.line.fill.background()
            txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.2), Inches(11), Inches(0.8))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = slide_data.get("title", "")
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = WHITE

            # 左侧图片
            img_keyword = slide_data.get("image_keyword", "finance")
            img = get_image(img_keyword)
            slide.shapes.add_picture(img, Inches(0.5), Inches(1.6), Inches(4.2), Inches(4.2))

            # 右侧要点
            points = slide_data.get("points", [])
            txBox = slide.shapes.add_textbox(Inches(5.2), Inches(1.6), Inches(7), Inches(4))
            tf = txBox.text_frame
            tf.word_wrap = True
            for i, point in enumerate(points):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = f"▸ {point}"
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(30,30,30)
                p.space_after = Pt(10)

            # 原文链接按钮（可点击）
            link_url = slide_data.get("link", "")
            if link_url:
                # 按钮形状
                btn = slide.shapes.add_shape(
                    1,  # 矩形
                    Inches(5.2), Inches(5.9), Inches(3), Inches(0.45)
                )
                btn.fill.solid()
                btn.fill.fore_color.rgb = PRIMARY_COLOR
                btn.line.fill.background()
                # 按钮文字
                tf_btn = btn.text_frame
                tf_btn.word_wrap = False
                p_btn = tf_btn.paragraphs[0]
                p_btn.text = "🔗 阅读原文"
                p_btn.font.size = Pt(14)
                p_btn.font.color.rgb = WHITE
                p_btn.alignment = PP_ALIGN.CENTER
                # 设置超链接
                btn.click_action.hyperlink.address = link_url

            # 可选图表
            chart_data = slide_data.get("chart")
            if chart_data and chart_data.get("categories") and chart_data.get("values"):
                try:
                    chart_frame = slide.shapes.add_chart(
                        XL_CHART_TYPE.BAR_CLUSTERED if chart_data.get("type")=="bar" else XL_CHART_TYPE.LINE,
                        Inches(0.5), Inches(5.9), Inches(4.2), Inches(1.3)
                    )
                    chart = chart_frame.chart
                    cat_data = CategoryChartData()
                    cat_data.categories = chart_data["categories"]
                    cat_data.add_series('', chart_data["values"])
                    chart.replace_data(cat_data)
                    chart.has_legend = False
                    chart.style = 2
                except Exception as e:
                    print(f"图表生成失败: {e}")

        # --- 总结页 ---
        elif stype == "summary":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(slide, PRIMARY_COLOR)
            txBox = slide.shapes.add_textbox(Inches(1.5), Inches(1.5), Inches(10), Inches(4))
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = "总结与展望"
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
            p.alignment = PP_ALIGN.CENTER
            summary_text = slide_data.get("text", "")
            p2 = tf.add_paragraph()
            p2.text = summary_text
            p2.font.size = Pt(24)
            p2.font.color.rgb = WHITE
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(30)

    ppt_filename = "daily_news.pptx"
    prs.save(ppt_filename)
    return ppt_filename

# ---------- 5. 发送到企业微信 ----------
def send_to_wecom(file_path, webhook):
    key = webhook.split('key=')[-1]
    upload_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={key}&type=file"
    with open(file_path, 'rb') as f:
        files = {'media': f}
        r = requests.post(upload_url, files=files)
    media_id = r.json()['media_id']
    data = {"msgtype": "file", "file": {"media_id": media_id}}
    requests.post(webhook, json=data)

# ---------- 主流程 ----------
if __name__ == '__main__':
    articles = fetch_news()
    if not articles:
        print("没有抓到新闻，退出")
        exit(1)
    today = datetime.now().strftime("%Y年%m月%d日")
    slides = generate_ppt_blueprint(articles, today)
    ppt_file = create_ppt_from_blueprint(slides)
    send_to_wecom(ppt_file, WECOM_WEBHOOK)
    print("推送完成")
