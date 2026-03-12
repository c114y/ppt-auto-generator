import streamlit as st
from openai import OpenAI
import docx
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.ns import qn
import concurrent.futures
import io
import re

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
        temperature=0.1 # 极低随机性，保障数据严谨、不被随意篡改或翻译
    )
    return clean_output(response.choices[0].message.content)

def generate_outline(text, custom_req, client, model_name):
    prompt = f"""
    结合以下论文内容，帮我做一份答辩PPT大纲及每个大纲对应的内容。预计15~18页。
    
    绝对不可违背的红线要求（违规将导致系统崩溃）：
    1. 必须且只能使用 ====PAGE_BREAK==== 作为每页PPT之间的分隔符（第一页开头不要加，放在每页中间）。
    2. 标题直接写具体内容，绝对不允许在标题前添加“第x页：”或类似字样。
    3. 严禁生成、提炼或修改图表数据！涉及所有财务数据展示，必须100%精确抓取原文真实原始数据，原样复制，严禁对数据进行任何形式的汇总、删减或概括。
    4. 所有年份、金额、百分比等数值，必须完整保留阿拉伯数字（如：2024、8295），严禁转换为中文大写（如二零二四）。
    5. 全文输出严禁掺杂任何额外的英文单词或生成环境的水印。
    6. 文本中严格禁止使用双星号(**)等任何Markdown格式符号进行排版。按分级标题和要点式内容生成，每点40-80字，用符号（-）分项。
    7. 客户附加要求优先执行：{custom_req}
    
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
    3. 必须以“【第X页演讲词】”作为每一段的开头，与大纲内容一一对应。
    4. 严禁使用双星号(**)等Markdown符号排版，严禁篡改原大纲中的阿拉伯数字。
    5. 严禁出现任何英文文本。
    
    PPT大纲：
    {outline}
    """
    return call_llm(prompt, client, model_name)

def build_ppt_file(outline, template_path="template.pptx"):
    """基于模板自动化排版并生成 PPTX 文件流"""
    try:
        prs = Presentation(template_path)
    except Exception as e:
        raise FileNotFoundError("找不到 template.pptx 文件，请确认已将母版上传至代码同一目录下！")
    
    pages = outline.split("====PAGE_BREAK====")
    
    for page in pages:
        if not page.strip(): continue
        
        # 调用模板中的「标题和内容」母版版式
        slide_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout) 
        
        lines = page.strip().split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        if not lines: continue
        
        # 正则清洗标题前缀
        raw_title = lines[0]
        clean_title = re.sub(r"^(第\d+页[：:]\s*)", "", raw_title)
        content_lines = lines[1:]
        
        # 写入标题并强制设定字体
        if slide.shapes.title:
            slide.shapes.title.text = clean_title
            for par in slide.shapes.title.text_frame.paragraphs:
                for run in par.runs:
                    set_font_style(run, 28)
                
        # 智能解析正文要点并写入模板占位符
        if len(slide.placeholders) > 1:
            body_shape = slide.placeholders[1]
            tf = body_shape.text_frame
            tf.clear() # 清除模板自带的提示占位符
            
            for line in content_lines:
                line = line.strip()
                if not line: continue
                
                p = tf.add_paragraph()
                
                # 剔除横杠并转换为原生层级缩进
                clean_line = line.replace("-", "").strip()
                p.text = clean_line
                
                # 注入空间层级排版逻辑
                p.level = 1 if line.startswith('-') else 0 
                p.space_after = Pt(14)  # 增加段后距，拉开呼吸感
                p.line_spacing = 1.3    # 1.3倍行距，避免文字拥挤
                
                for run in p.runs:
                    set_font_style(run, 18)
                
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# ================= Streamlit UI 界面 =================
st.set_page_config(page_title="自动化答辩PPT系统", layout="wide")

# 初始化会话缓存
if 'speech_text' not in st.session_state:
    st.session_state.speech_text = None
if 'ppt_io' not in st.session_state:
    st.session_state.ppt_io = None

# --- 侧边栏 API 配置 ---
st.sidebar.header("⚙️ API 配置")
api_key = st.sidebar.text_input("API Key", type="password")
base_url = st.sidebar.text_input("接口地址 (Base URL)", value="https://api.packyapi.com/v1") 
model_name = st.sidebar.text_input("模型名称 (Model)", value="gemini-3.1-pro-preview")

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
            client = OpenAI(api_key=api_key, base_url=base_url)
            
            with st.spinner("正在提取论文数据..."):
                doc_text = extract_text_from_docx(uploaded_file)
            
            st.info("第一步：正在精准提取源文档并生成 PPT 大纲...")
            outline = generate_outline(doc_text, custom_prompt, client, model_name)
            
            with st.expander("👀 预览生成的大纲 (点击展开)"):
                st.write(outline)
            
            st.info("第二步：正在多线程同步生成演讲稿与挂载模板排版...")
            
            with concurrent.futures.ThreadPoolExecutor() as executor:
                future_speech = executor.submit(generate_speech, outline, client, model_name)
                # 代码固定读取同目录下的 template.pptx
                future_ppt = executor.submit(build_ppt_file, outline, "template.pptx")
                
                st.session_state.speech_text = future_speech.result()
                st.session_state.ppt_io = future_ppt.result()
                
            st.success("🎉 全部生成成功！请点击下方按钮下载。")
                
        except Exception as e:
            st.error(f"生成过程中出现错误，请检查配置或文档内容: {e}")

if st.session_state.speech_text and st.session_state.ppt_io:
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📥 下载配套演讲稿 (.txt)", 
            data=st.session_state.speech_text, 
            file_name="答辩演讲稿.txt", 
            mime="text/plain"
        )
    with col2:
        st.download_button(
            label="📥 下载排版后PPT (.pptx)", 
            data=st.session_state.ppt_io, 
            file_name="自动化排版答辩.pptx", 
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )