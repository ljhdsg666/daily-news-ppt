import os
import io
import requests
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from openai import OpenAI
from datetime import datetime
import feedparser
from PIL import Image

# 配置
DEEPSEEK_API_KEY = os.environ['DEEPSEEK_API_KEY']
WECOM_WEBHOOK = os.environ['WECOM_WEBHOOK']

client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

# ---------- 1. 抓取新闻 ----------
def fetch_news():
    """Google 新闻 RSS 抓取 US News 关于 Wall Street 的文章"""
    query = "wall street site:usnews.com"
    url = f"https://news.google.com/rss/search?q={requests.utils.quote(query)}&hl=en-US&gl=US&ceid=US:en"
    feed = feedparser.parse(url)
    articles = []
    for entry in feed.entries[:10]:
        articles.append({
            "title": entry.title,
            "link": entry.link,
        })
    if not articles:
        print("未抓取到新闻")
    return articles

# ---------- 2. AI 生成详细摘要 ----------
def generate_detailed_summary(articles):
    """为每条新闻生成包含要点和一句话分析的详细描述"""
    enhanced = []
    for idx, art in enumerate(articles):
        prompt = f"""你是一个资深财经编辑，请根据以下新闻标题，用中文写一段100~150字的摘要，风格专业、有洞察。包含：
- 一句话概括核心事件
- 一点简要分析或市场影响
标题：{art['title']}
"""
        try:
            resp = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
            )
            summary = resp.choices[0].message.content
        except Exception as e:
            print(f"AI 摘要失败: {e}")
            summary = art['title']
        enhanced.append({
            "title": art['title'],
            "summary": summary.strip(),
            "link": art['link']
        })
    return enhanced

# ---------- 3. 获取免版权图片 ----------
def get_image_for_keyword(keyword):
    """从 Unsplash 免费图库搜索并返回图片的二进制数据（无API Key）"""
    search_url = f"https://source.unsplash.com/800x600/?{requests.utils.quote(keyword)}"
    try:
        resp = requests.get(search_url, timeout=10)
        if resp.status_code == 200:
            img = Image.open(io.BytesIO(resp.content))
            # 转换为 PNG 存入内存
            img_io = io.BytesIO()
            img.save(img_io, format='PNG')
            img_io.seek(0)
            return img_io
    except Exception as e:
        print(f"图片下载失败 ({keyword}): {e}")
    return None

# ---------- 4. 生成专业PPT ----------
def create_professional_ppt(articles_enhanced, date_str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 宽屏
    prs.slide_height = Inches(7.5)

    # --- 封面页 ---
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白页
    # 尝试加背景图
    bg_img = get_image_for_keyword("wall street finance")
    if bg_img:
        # 将背景图加到幻灯片
        from pptx.util import Emu
        slide.shapes.add_picture(bg_img, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    # 半透明黑色遮罩
    left = top = Inches(0)
    width = prs.slide_width
    height = prs.slide_height
    shape = slide.shapes.add_shape(
        1, left, top, width, height  # 1 = MSO_SHAPE.RECTANGLE
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 0, 0)
    shape.fill.fore_color.brightness = -0.4  # 半透明效果
    shape.line.fill.background()
    # 标题文字
    txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "每日华尔街简报"
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    # 日期副标题
    txBox2 = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(11), Inches(1))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = date_str
    p2.font.size = Pt(28)
    p2.font.color.rgb = RGBColor(200, 200, 200)
    p2.alignment = PP_ALIGN.CENTER

    # --- 每条新闻单独一页 ---
    for art in articles_enhanced:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 左侧图片区域
        img_keyword = " ".join(art['title'].split()[:3])  # 取前3个词做关键词
        img = get_image_for_keyword(img_keyword)
        if img:
            pic = slide.shapes.add_picture(img, Inches(0.5), Inches(1.2), Inches(4.5), Inches(4.5))
        # 右侧文本框
        left = Inches(5.5)
        top = Inches(1.2)
        width = Inches(7)
        height = Inches(5)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        # 新闻标题
        p = tf.paragraphs[0]
        p.text = art['title']
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
        p.space_after = Pt(12)
        # 摘要
        p2 = tf.add_paragraph()
        p2.text = art['summary']
        p2.font.size = Pt(18)
        p2.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        p2.space_after = Pt(8)
        # 来源链接（可点击不现实，写个小字）
        p3 = tf.add_paragraph()
        p3.text = "来源: US News"
        p3.font.size = Pt(12)
        p3.font.color.rgb = RGBColor(150, 150, 150)

    # --- 结尾页 ---
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    txBox = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(9), Inches(2))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "以上信息由AI自动整理，仅供参考\n不构成任何投资建议"
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(100, 100, 100)
    p.alignment = PP_ALIGN.CENTER

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
    data = {
        "msgtype": "file",
        "file": {"media_id": media_id}
    }
    requests.post(webhook, json=data)

# ---------- 主流程 ----------
if __name__ == '__main__':
    articles = fetch_news()
    if not articles:
        print("没有抓到新闻，退出")
        exit(1)
    # 生成详细摘要
    articles_enhanced = generate_detailed_summary(articles)
    today = datetime.now().strftime("%Y年%m月%d日")
    ppt_file = create_professional_ppt(articles_enhanced, today)
    send_to_wecom(ppt_file, WECOM_WEBHOOK)
    print("推送完成")
