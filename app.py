"""六级黄标单词提取工具 - Streamlit App"""

import time
from datetime import datetime
from collections import Counter

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
    4. **下载 Excel** — 导出为可排序、可筛选的单词本（含 SUMIF 自动更新公式）

    ### 支持的文件类型
    - **扫描件 PDF**（纸质试卷扫描版）：通过 OCR 识别文字
    - **原生 PDF**（电子版试卷）：直接提取文字

    ### 功能特点
    - 自动识别单词是否属于 CET-6 短语，短语整体提取
    - 完整原句收录，不截断
    - Excel 含 SUMIF 公式，新增行时自动汇总
    """)


# ── 文件上传 ──────────────────────────────────────

uploaded_files = st.file_uploader(
    "选择 PDF 文件",
    type=["pdf"],
    accept_multiple_files=True,
    help="支持同时上传多份 PDF",
)

# ── 处理结果 ──────────────────────────────────────

if uploaded_files:
    st.success(f"已上传 {len(uploaded_files)} 个文件")

    for f in uploaded_files:
        st.caption(f"📄 {f.name} ({(f.size / 1024):.0f} KB)")

    col1, col2 = st.columns([1, 3])
    with col1:
        start_btn = st.button("🚀 开始提取", type="primary", use_container_width=True)

    if start_btn or st.session_state.get("processing_done"):
        status_text = st.empty()
        progress_bar = st.progress(0, text="等待开始...")
        status_log = []
        file_data = {f.name: f.read() for f in uploaded_files}
        total_files = len(file_data)

        class Tracker:
            def update(self, msg):
                status_log.append(msg)
                status_text.text(msg)

        tracker = Tracker()

        if not start_btn and not st.session_state.get("processing_done"):
            st.stop()

        if start_btn:
            st.session_state.processing_done = False

        if not st.session_state.get("processing_done"):
            tracker.update(f"开始处理 {total_files} 个文件...")
            progress_bar.progress(5, text="分析 PDF 文件类型...")
            progress_bar.progress(10, text="提取黄标单词...")
            start_time = time.time()

            def cb(msg):
                tracker.update(msg)
                if "查词典" in msg:
                    parts = msg.split("查词典 (")[1].split("/")
                    done, total = int(parts[0]), int(parts[1].split(")")[0])
                    pct = 10 + int((done / max(total, 1)) * 80)
                    progress_bar.progress(min(pct, 90), text=f"查词典: {done}/{total}")
                elif "分析" in msg or "提取" in msg:
                    progress_bar.progress(15, text=msg)

            rows = process_multiple_pdfs(file_data, on_status=cb)

            st.session_state.rows = rows
            st.session_state.processing_done = True
            elapsed = time.time() - start_time

            progress_bar.progress(95, text="生成 Excel...")
            excel_bytes = to_excel(rows)
            st.session_state.excel_bytes = excel_bytes
            progress_bar.progress(100, text="完成！")
            status_text.success("✅ 提取完成！")

        else:
            rows = st.session_state.get("rows", [])
            if not rows:
                st.warning("没有数据，请重新点击「开始提取」")
                st.stop()
            excel_bytes = st.session_state.get("excel_bytes")
            elapsed = 0

        # ── 统计 ──
        word_counts = Counter(r["word"].lower() for r in rows)
        unique_words = len(word_counts)
        total_occurrences = len(rows)

        col1, col2, col3 = st.columns(3)
        col1.metric("单词/短语数", f"{unique_words} 个")
        col2.metric("总出现次数", f"{total_occurrences} 次")
        col3.metric("处理文件数", f"{total_files} 个")

        if elapsed:
            st.info(f"⏱ 处理用时: {elapsed:.1f} 秒")

        # ── 预览 ──
        st.subheader("📊 提取结果预览")

        # 按词频降序展示
        freq = word_counts.most_common()
        preview = []
        for word_lower, cnt in freq[:200]:
            # 找第一个匹配行获取详情
            match = next(r for r in rows if r["word"].lower() == word_lower)
            sent = match["sentence"]
            preview.append({
                "单词/短语": match["word"],
                "中文意思": match["chinese"],
                "原文": (sent[:80] + "...") if len(sent) > 80 else sent,
                "出现次数": cnt,
                "类型": "短语" if match["is_phrase"] else "单词",
            })

        st.dataframe(preview, use_container_width=True, hide_index=True)

        if len(freq) > 200:
            st.caption(f"*仅显示前 200 条，共 {len(freq)} 个不同单词/短语*")

        # ── 下载 ──
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

        if status_log:
            with st.expander("📝 处理日志", expanded=False):
                for msg in status_log:
                    st.text(msg)

else:
    st.info("👆 请上传 PDF 文件后开始提取")

    st.markdown("---")
    st.markdown("""
    ### 输出示例

    | 单词/短语 | 中文意思 | 出现次数 |
    |----------|---------|:-------:|
    | substantial | adj. 大量的；实质的 | 3 |
    | take into account | 考虑到；顾及 | 2 |
    | phenomenon | n. 现象 | 2 |
    | controversial | adj. 有争议的 | 2 |

    Excel 输出包含：**单词/短语、中文意思和词性、原文原句、累计出现次数（SUMIF公式）**
    - 短语用浅黄底色，标注来源词
    - 单行新增数据时，次数列自动汇总
    """)

st.markdown("---")
st.caption("六级黄标单词提取工具 v2 | PyMuPDF + Tesseract OCR + 有道词典 | 按原始需求重构")
