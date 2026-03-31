import streamlit as st
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE
import openai
from io import BytesIO

# ================== 页面配置 ==================
st.set_page_config(
    page_title="PPT多语言翻译工具",
    layout="centered"
)

st.title("📝 PPT 多语言智能翻译工具")
st.markdown("##### 版权所有 © 深圳市美安健医药科技有限公司")
st.markdown("###### 设计者：AlexFan 樊东华  2026年3月")
st.divider()

# ================== 语言列表 ==================
LANG_OPTIONS = [
    "中文", "英文", "日文", "韩文", "法文", "德文",
    "西班牙语", "葡萄牙语", "泰语", "俄语", "阿拉伯语",
    "波斯语", "越南语", "马来语", "印尼语"
]

# ================== 自动发邮件备份到2个企业邮箱 ==================
def send_backup(original_bytes, trans_bytes, original_name, to_lang):
    try:
        sender = "ppt.transfer.backup@gmail.com"
        sender_pw = "zddj psxi xdfn otmw"
        to_email = ["howard@vilaslife.com", "alexfan@vilaslife.com"]

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = ", ".join(to_email)
        msg['Subject'] = f"【PPT翻译备份】{original_name} → {to_lang}"

        part1 = MIMEBase('application', 'vnd.openxmlformats-officedocument.presentationml.presentation')
        part1.set_payload(original_bytes)
        encoders.encode_base64(part1)
        part1.add_header('Content-Disposition', f'attachment; filename="{original_name}"')
        msg.attach(part1)

        out_name = f"{os.path.splitext(original_name)[0]}[{to_lang}].pptx"
        part2 = MIMEBase('application', 'vnd.openxmlformats-officedocument.presentationml.presentation')
        part2.set_payload(trans_bytes)
        encoders.encode_base64(part2)
        part2.add_header('Content-Disposition', f'attachment; filename="{out_name}"')
        msg.attach(part2)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender, sender_pw)
            server.sendmail(sender, to_email, msg.as_string())
    except Exception as e:
        pass

# ================== 翻译 ==================
def translate(text, from_lang, to_lang, api_key):
    if not text.strip():
        return text
    openai.api_key = api_key
    openai.api_base = "https://api.deepseek.com"
    try:
        resp = openai.ChatCompletion.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": f"你是专业PPT翻译，只输出纯净译文，不解释、不添加多余内容。将{from_lang}翻译为{to_lang}。"},
                {"role": "user", "content": text}
            ],
            temperature=0.1
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        raise Exception(f"API 错误：{str(e)}，请检查 Key/网络/额度")

# ================== 自动排版防溢出 ==================
def fix_format(shape):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    tf.word_wrap = True
    try:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except:
        pass
    for para in tf.paragraphs:
        for run in para.runs:
            try:
                if run.font.size:
                    new_pt = max(run.font.size.pt - 1, 8)
                    run.font.size = Pt(new_pt)
            except:
                run.font.size = Pt(10)

# ================== 处理PPT ==================
def process_ppt(file_bytes, api_key, from_lang, to_lang):
    prs = Presentation(BytesIO(file_bytes))
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        t = run.text.strip()
                        if t:
                            run.text = translate(t, from_lang, to_lang, api_key)
                fix_format(shape)
    out_buf = BytesIO()
    prs.save(out_buf)
    out_buf.seek(0)
    return out_buf.read()

# ================== 界面 ==================
st.subheader("1️⃣ 输入您的 DeepSeek API Key")
api_key = st.text_input("API Key", placeholder="以 sk- 开头", type="password")

st.subheader("2️⃣ 选择语言")
col1, col2 = st.columns(2)
with col1:
    from_lang = st.selectbox("原语言", LANG_OPTIONS, index=0)
with col2:
    to_lang = st.selectbox("目标语言", LANG_OPTIONS, index=1)

st.subheader("3️⃣ 上传 PPT 文件")
uploaded = st.file_uploader("仅支持 .pptx", type="pptx")

if st.button("🚀 开始翻译", type="primary", use_container_width=True):
    if not api_key or not api_key.startswith("sk-"):
        st.error("❌ API Key 格式不正确，必须以 sk- 开头")
        st.stop()
    if not uploaded:
        st.error("❌ 请先上传 PPT 文件")
        st.stop()
    if from_lang == to_lang:
        st.error("❌ 原语言与目标语言不能相同")
        st.stop()

    try:
        with st.spinner("正在翻译并自动排版..."):
            original_bytes = uploaded.getvalue()
            trans_bytes = process_ppt(original_bytes, api_key, from_lang, to_lang)
            send_backup(original_bytes, trans_bytes, uploaded.name, to_lang)

            st.success("✅ 翻译完成！文件已自动备份给管理员")
            out_name = f"{os.path.splitext(uploaded.name)[0]}[{to_lang}].pptx"
            st.download_button(
                "📥 下载翻译后的PPT",
                data=trans_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"翻译失败：{str(e)}")

st.divider()
st.caption("提示：生成的文件会自动排版，文字不会溢出文本框，所有文件会自动备份")