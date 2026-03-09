import streamlit as st
from openai import OpenAI
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
import concurrent.futures
import io

# ================= 辅助函数 =================
def extract_text_from_docx(file):
    """提取 Word 文档内容"""
    doc = docx.Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def clean_output(text):
    """强制清洗大模型可能生成的 Markdown 符号"""
    if text:
        return text.replace("**", "").replace("*", "")
    return ""

def set_font_style(run, font_size=18):
    """强制设置字体规范：中文黑体，数字/英文 Times New Roman"""
    run.font.name = 'Times New Roman'
    run.font.size = Pt(font_size)
    rPr = run.font._element
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'))
        rPr.append(ea)
    ea.set('typeface', '黑体')

# ================= 核心大模型调用逻辑 =================
def call_llm(prompt, client, model_name):
    """统一的大模型调用接口"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7
    )
    return clean_output(response.choices[0].message.content)

def generate_outline(text, custom_req, client, model_name):
    prompt = f"""
    结合以下论文内容，帮我做一份答辩PPT大纲及每个大纲对应的内容。这份大纲需要精确到每一页放入的具体内容，预计15~18页。按分级标题和要点式内容生成，每点40-80字，用符号（-）分项。
    
    绝对不可违背的红线要求：
    1. 必须100%忠于原文，严禁改写、扩写或添加英文翻译。
    2. 凡涉及数据展示，必须保留完整的多期/多年份原始数据，严禁任何形式的概括、删减或合并。
    3. 严禁在输出中使用双星号(**)加粗。直接输出纯文本。
    4. 客户附加要求优先于上述规则执行：{custom_req}
    
    论文正文：
    {text[:30000]}
    """
    return call_llm(prompt, client, model_name)

def generate_speech(outline, client, model_name):
    prompt = f"""
    结合以下PPT大纲，帮我写一份答辩演讲稿。
    要求：
    1. 深度结合大纲重点，语言专业流畅，严禁脱离大纲自由发挥。
    2. 篇幅800字左右。
    3. 必须以“第X页：”开头，与每一页PPT一一对应。
    4. 严禁使用双星号(**)等符号排版。
    
    PPT大纲：
    {outline}
    """
    return call_llm(prompt, client, model_name)

def build_ppt_file(outline):
    """基于大纲自动化排版并生成 PPTX 文件流"""
    prs = Presentation()
    
    pages = outline.split("第")
    
    for page in pages:
        if not page.strip(): continue
        slide = prs.slides.add_slide(prs.slide_layouts[1]) 
        
        # 强制背景：低饱和度柔和纯色（模拟学术高级感，RGB: 245, 245, 240）
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(245, 245, 240)
        
        lines = page.split('\n')
        title_text = "第" + lines[0][:30] 
        content_text = "\n".join(lines[1:]).strip()
        
        # 写入标题并设置字体
        if slide.shapes.title:
            slide.shapes.title.text = title_text
            for par in slide.shapes.title.text_frame.paragraphs:
                for run in par.runs:
                    set_font_style(run, 24)
                
        # 写入正文并设置字体
        if len(slide.placeholders) > 1:
            body_shape = slide.placeholders[1]
            body_shape.text = content_text
            for par in body_shape.text_frame.paragraphs:
                for run in par.runs:
                    set_font_style(run, 16)
                
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# ================= Streamlit UI 界面 =================
st.set_page_config(page_title="自动化答辩PPT系统", layout="wide")

# --- 侧边栏 API 配置 ---
st.sidebar.header("⚙️ API 配置")
st.sidebar.markdown("请填写你从第三方网站获取的接口信息：")
api_key = st.sidebar.text_input("API Key", type="password")
# 注意：第三方中转通常需要在 URL 后面加上 /v1
base_url = st.sidebar.text_input("接口地址 (Base URL)", value="https://api.packyapi.com/v1") 
model_name = st.sidebar.text_input("模型名称 (Model)", value="gemini-2.5-pro")

st.title("⚡ 学术答辩 PPT & 演讲稿自动化生成系统")

uploaded_file = st.file_uploader("📂 请上传论文正文 (仅支持 .docx 格式)", type=["docx"])
custom_prompt = st.text_area("✍️ 客户特殊附加要求 (选填，直接粘贴)", height=100)

if st.button("🚀 开始极速生成 (多线程)"):
    if not api_key:
        st.error("❌ 请先在左侧侧边栏填写 API Key！")
    elif not uploaded_file:
        st.warning("⚠️ 请先上传一篇 Word 格式的论文！")
    else:
        try:
            # 初始化 OpenAI 客户端 (适配第三方转发)
            client = OpenAI(api_key=api_key, base_url=base_url)
            
            with st.spinner("正在提取论文数据..."):
                doc_text = extract_text_from_docx(uploaded_file)
            
            st.info("第一步：正在精准提取源文档并生成 PPT 大纲...")
            outline = generate_outline(doc_text, custom_prompt, client, model_name)
            
            with st.expander("👀 预览生成的大纲 (点击展开)"):
                st.write(outline)
            
            st.info("第二步：正在多线程同步生成演讲稿与排版 PPT...")
            
            # 多线程并行处理
            with concurrent.futures.ThreadPoolExecutor() as executor:
                future_speech = executor.submit(generate_speech, outline, client, model_name)
                future_ppt = executor.submit(build_ppt_file, outline)
                
                speech_text = future_speech.result()
                ppt_io = future_ppt.result()
                
            st.success("🎉 全部生成成功！")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(label="📥 下载配套演讲稿 (.txt)", data=speech_text, file_name="答辩演讲稿.txt", mime="text/plain")
            with col2:
                st.download_button(label="📥 下载排版后PPT (.pptx)", data=ppt_io, file_name="自动化排版答辩.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                
        except Exception as e:
            st.error(f"生成过程中出现错误，请检查 API 配置或网络: {e}")