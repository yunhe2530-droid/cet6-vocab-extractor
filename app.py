"""六级黄标单词提取工具 - Streamlit App"""

import io
import time
from datetime import datetime

import streamlit as st

from extractor import process_multiple_pdfs, to_excel

# ── 页面配置 ──────────────────────────────────────

st.set_page_config(
    page_title="六级黄标单词提取",
    page_icon="📖",
    layout="wide",
)

st.title("📖 六级黄标单词提取工具")
st.markdown("上传六级真题 PDF（支持多个），自动提取黄色标记的单词并生成 Excel 单词本。")

# ── 说明区 ────────────────────────────────────────

with st.expander("📋 使用说明", expanded=False):
    st.markdown("""
    ### 使用步骤
    1. **上传 PDF** — 支持同时上传多份六级真题 PDF（扫描件或原生文本均可）
    2. **点击「开始提取」** — 系统自动检测黄色标记区域，提取单词
    3. **预览结果** — 在线查看提取的单词列表
    4. **下载 Excel** — 导出为可排序、可筛选的单词本

    ### 支持的文件类型
    - **扫描件 PDF**（纸质试卷扫描版）：通过 OCR 识别文字
    - **原生 PDF**（电子版试卷）：直接提取文字

    ### 注意事项
    - 黄色标记必须是 PDF 内的**绘图矩形标记**（手写笔记软件生成的那种）
    - 处理扫描件需要安装 Tesseract OCR（已在 `C:\\Program Files\\Tesseract-OCR`）
    - 词典查询需要网络连接，每个单词会查有道词典
    - 处理时间取决于 PDF 页数和单词数量
    """)

# ── 文件上传 ──────────────────────────────────────

uploaded_files = st.file_uploader(
    "选择 PDF 文件",
    type=["pdf"],
    accept_multiple_files=True,
    help="支持同时上传多份 PDF"
)

# ── 处理结果 ──────────────────────────────────────

if uploaded_files:
    st.success(f"已上传 {len(uploaded_files)} 个文件")

    # 文件信息
    file_info = [(f.name, f"{(f.size / 1024):.0f} KB") for f in uploaded_files]
    for name, size in file_info:
        st.caption(f"📄 {name} ({size})")

    # 提取按钮
    col1, col2 = st.columns([1, 3])
    with col1:
        start_btn = st.button("🚀 开始提取", type="primary", use_container_width=True)

    if start_btn or st.session_state.get("processing_done"):
        status_area = st.container()
        progress_bar = st.progress(0, text="等待开始...")
        status_text = st.empty()

        # 状态回调
        status_log = []
        file_data = {}
        for f in uploaded_files:
            file_data[f.name] = f.read()

        total_files = len(file_data)

        # 进度跟踪
        class ProgressTracker:
            def __init__(self):
                self.step = 0

            def update(self, msg):
                status_log.append(msg)
                status_text.text(msg)

        tracker = ProgressTracker()

        if not start_btn and not st.session_state.get("processing_done"):
            st.stop()

        if start_btn:
            st.session_state.processing_done = False

        if not st.session_state.get("processing_done"):
            tracker.update(f"开始处理 {total_files} 个文件...")

            # 阶段1: PDF分析
            progress_bar.progress(5, text="分析 PDF 文件类型...")

            # 阶段2: 提取
            progress_bar.progress(10, text="提取黄标单词...")

            start_time = time.time()

            def status_callback(msg):
                tracker.update(msg)
                # 根据消息更新进度
                if "查词典" in msg:
                    parts = msg.split("查词典 (")[1].split("/")
                    done = int(parts[0])
                    total_dict = int(parts[1].split(")")[0])
                    overall_progress = 10 + int((done / max(total_dict, 1)) * 80)
                    progress_bar.progress(
                        min(overall_progress, 90),
                        text=f"查词典: {done}/{total_dict}"
                    )
                elif "分析" in msg or "提取" in msg:
                    progress_bar.progress(15, text=msg)

            words_data = process_multiple_pdfs(file_data, on_status=status_callback)

            st.session_state.words_data = words_data
            st.session_state.processing_done = True
            elapsed = time.time() - start_time

            progress_bar.progress(95, text="生成 Excel...")

            # 生成 Excel
            excel_bytes = to_excel(words_data)
            st.session_state.excel_bytes = excel_bytes

            progress_bar.progress(100, text="完成！")
            status_text.success("✅ 提取完成！")

        else:
            # 从 session_state 恢复
            words_data = st.session_state.get("words_data", [])
            if not words_data:
                st.warning("没有数据，请重新点击「开始提取」")
                st.stop()
            excel_bytes = st.session_state.get("excel_bytes")
            elapsed = 0

        # ── 结果显示 ──────────────────────────────

        # 统计
        total_words = len(words_data)
        total_count = sum(w["total_count"] for w in words_data)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("单词总数", f"{total_words} 个")
        with col2:
            st.metric("总出现次数", f"{total_count} 次")
        with col3:
            st.metric("处理文件数", f"{total_files} 个")

        if elapsed > 0:
            st.info(f"⏱ 处理用时: {elapsed:.1f} 秒")

        # 预览表格
        st.subheader("📊 提取结果预览")
        preview_data = []
        for w in words_data[:200]:  # 预览前200条
            sentences = w["sentences"][0] if w["sentences"] else ""
            if len(sentences) > 80:
                sentences = sentences[:80] + "..."
            preview_data.append({
                "单词/短语": w["word"],
                "中文意思": w["chinese"],
                "原文": sentences,
                "出现次数": w["total_count"],
            })

        st.dataframe(preview_data, use_container_width=True, hide_index=True)

        if total_words > 200:
            st.caption(f"*仅显示前 200 条，共 {total_words} 条，下载 Excel 可查看全部*")

        # Excel 下载
        st.subheader("⬇️ 导出")
        if excel_bytes:
            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="📥 下载 Excel 单词本",
                data=excel_bytes,
                file_name=f"六级标黄单词本_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        # 原始处理日志（折叠）
        if status_log:
            with st.expander("📝 处理日志", expanded=False):
                for msg in status_log:
                    st.text(msg)

else:
    # 空状态
    st.info("👆 请上传 PDF 文件后开始提取")

    # 样例预览
    st.markdown("---")
    st.markdown("""
    ### 输出示例

    | 单词/短语 | 中文意思 | 出现次数 |
    |----------|---------|:-------:|
    | substantial | adj. 大量的；实质的 | 3 |
    | controversial | adj. 有争议的 | 2 |
    | take into account | 考虑到；顾及 | 2 |
    | phenomenon | n. 现象 | 2 |
    | perspective | n. 视角；观点 | 1 |

    Excel 输出包含：**单词/短语、中文意思和词性、原文原句、累计出现次数**
    """)

st.markdown("---")
st.caption("六级黄标单词提取工具 | PyMuPDF + Tesseract OCR + 有道词典")
