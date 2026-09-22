# -*- coding: utf-8 -*-
"""
مساعد حالة العميل — نسخة ويب (Streamlit)
==========================================
نفس فكرة البرنامج المكتبي، بس شغّال من المتصفح عشان أي حد في الفريق (في أي محافظة) يستخدمه
من غير ما يثبّت بايثون: يرفع ملف التقرير (PDF أو ASPX)، وياخد إكسيل جاهز، أو يفتح مساعد
التصفية وهو متعبّي بالأقساط تلقائي.

التشغيل محليًا:  streamlit run app.py
الملفات المطلوبة جنب app.py: core.py + logo.png + مساعد_تصفية_العملاء_من_برنامج_المحصل.html
"""
import base64
import io
from pathlib import Path

import streamlit as st

import core

HERE = Path(__file__).parent
LIQ_HTML_PATH = HERE / "مساعد_تصفية_العملاء_من_برنامج_المحصل.html"


def find_logo():
    """
    بيدوّر على أي ملف اسمه logo بأي حروف كبيرة/صغيرة وأي امتداد صورة شائع،
    عشان يشتغل حتى لو الملف اترفع باسم Logo.PNG أو logo.jpg مثلًا.
    """
    for p in HERE.iterdir():
        if p.is_file() and p.stem.lower() == "logo" and p.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"):
            return p
    return None


LOGO_PATH = find_logo()

BRAND_RED = "#E30613"
BRAND_RED_DARK = "#B8050F"
BRAND_PINK = "#FBEAEC"


def open_in_new_tab_button(label, html_content, key):
    """
    زرار حقيقي بيفتح صفحة HTML (مبنية في الذاكرة) في تبويب جديد في متصفح المستخدم،
    بدل ما تتحط جوه إطار صغير في نفس الصفحة. المتصفحات بتمنع فتح تبويب تلقائي من غير
    ضغطة مستخدم، فالفتح لازم يحصل داخل onclick نفسه (مش بعد استدعاء بايثون).
    """
    b64 = base64.b64encode(html_content.encode("utf-8")).decode("ascii")
    st.components.v1.html(
        f"""
        <button id="{key}" style="
            background:linear-gradient(135deg,{BRAND_RED},{BRAND_RED_DARK}); color:white; border:none;
            border-radius:10px; padding:12px 18px; font-size:15px; font-weight:bold; cursor:pointer;
            width:100%; box-shadow:0 2px 6px rgba(227,6,19,0.35); font-family:inherit;
        ">{label}</button>
        <script>
        document.getElementById("{key}").onclick = function () {{
            var bytes = Uint8Array.from(atob("{b64}"), c => c.charCodeAt(0));
            var html = new TextDecoder("utf-8").decode(bytes);
            var blob = new Blob([html], {{type: "text/html"}});
            var url = URL.createObjectURL(blob);
            window.open(url, "_blank");
        }};
        </script>
        """,
        height=58,
    )


st.set_page_config(page_title="مساعد حالة العميل | Caritas Egypt", page_icon="📄", layout="centered")

# ---------- تنسيق عام: اتجاه، ألوان، أيقونات ----------
st.markdown(
    f"""
    <style>
    /* RTL قوي: بنستهدف حاويات Streamlit الحقيقية بدل الاعتماد على أسماء classes بتتغيّر بين النسخ */
    html, body,
    [data-testid="stAppViewContainer"], [data-testid="stMain"], [data-testid="stVerticalBlock"],
    [data-testid="stHorizontalBlock"], [data-testid="stMarkdownContainer"],
    [data-testid="stExpander"], [data-testid="stExpanderDetails"],
    .block-container {{
        direction: rtl !important;
    }}
    [data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li {{
        text-align: right !important;
    }}
    [data-testid="stMarkdownContainer"] ol, [data-testid="stMarkdownContainer"] ul {{
        direction: rtl !important; text-align: right !important;
        padding-right: 1.4em !important; padding-left: 0 !important; margin-right: 0 !important;
    }}

    .block-container {{ padding-top: 1.2rem; }}

    .brand-header {{ text-align: center; margin-bottom: 6px; }}
    .brand-header img {{ max-width: 190px; height: auto; margin-bottom: 6px; }}
    .brand-header h1 {{ margin: 4px 0 0; font-size: 1.5rem; color: {BRAND_RED_DARK}; }}
    .brand-header p {{ margin: 2px 0 0; color: #6b6b6b; font-size: 0.92rem; }}

    div[data-testid="stMetric"] {{
        background: {BRAND_PINK}; border-radius: 12px; padding: 10px 6px;
        border: 1px solid #f3d3d7;
    }}
    div[data-testid="stMetricLabel"] {{ color: {BRAND_RED_DARK}; }}

    .stTabs [data-baseweb="tab"] {{ font-weight: 700; font-size: 1rem; }}
    .stTabs [aria-selected="true"] {{ color: {BRAND_RED_DARK} !important; }}

    .warn-banner {{
        background: #FFF6E5; border: 1px solid #F2C866; color: #7A5200;
        border-radius: 10px; padding: 10px 16px; margin: 16px 0; font-weight: 600; text-align: center;
    }}
    .app-footer {{
        text-align: center; color: #9a9a9a; font-size: 0.82rem; margin-top: 34px;
        padding-top: 14px; border-top: 1px solid #eee;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- الهيدر: اللوجو (متوسّط وأصغر) فوق العنوان ----------
if LOGO_PATH is not None:
    logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    ext = LOGO_PATH.suffix.lstrip(".").lower()
    mime = "jpeg" if ext in ("jpg", "jpeg") else ext
    st.markdown(
        f"""
        <div style="text-align:center; margin-bottom:4px;">
            <img src="data:image/{mime};base64,{logo_b64}"
                 style="width:130px; height:auto; max-width:100%; display:inline-block; object-fit:contain;">
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    files_here = ", ".join(sorted(p.name for p in HERE.iterdir() if p.is_file())) or "(مفيش ملفات خالص)"
    st.caption(f"⚠️ مفيش ملف اسمه logo (png/jpg) في فولدر البرنامج.\nالملفات الموجودة: {files_here}")

st.markdown(
    """
    <div class="brand-header">
        <h1>📄 مساعد حالة العميل</h1>
        <p>تحويل تقرير "حالة العميل" لإكسيل، وتجهيز مساعد التصفية تلقائيًا</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="warn-banner">⚠️ النظام تحت الاختبار والتعديل.. يرجى مراجعة النتائج بعناية قبل الاعتماد عليها.</div>',
    unsafe_allow_html=True,
)

# ---------- شرح طريقة الاستخدام ----------
with st.expander("ℹ️ طريقة استخدام البرنامج", expanded=False):
    steps = [
        ("📤", "ارفع ملف التقرير (PDF أو ASPX) اللي حفظته من صفحة \"حالة العميل\"."),
        ("👀", "هتظهر ملخص سريع ببيانات العميل، وتقدر تفتح تفاصيل العميل والتمويل كاملة لو حبيت."),
        ("⬇️", "من تبويب \"تحويل لإكسيل\": دوس تنزيل، وهتاخد ملف Excel فيه بيانات العميل، التمويل، الضامنين، والأقساط."),
        ("🧮", "من تبويب \"مساعد التصفية\": دوس الزرار، وهتتفتحلك صفحة حساب التصفية في تبويب جديد، بالأقساط متعبّية أوتوماتيك."),
        ("🔁", "عايز تشتغل على عميل تاني؟ ارفع ملفه من نفس الصفحة وكرر الخطوات."),
    ]
    rows = "".join(
        f"""
        <div style="display:flex; flex-direction:row-reverse; align-items:flex-start; gap:10px; margin:8px 0;">
            <div style="background:{BRAND_PINK}; color:{BRAND_RED_DARK}; font-weight:700; min-width:28px;
                        height:28px; border-radius:50%; display:flex; align-items:center; justify-content:center;">{i}</div>
            <div style="flex:1; text-align:right;">{icon} {text}</div>
        </div>"""
        for i, (icon, text) in enumerate(steps, 1)
    )
    st.markdown(f'<div style="direction:rtl;">{rows}</div>', unsafe_allow_html=True)

st.divider()

# ---------- رفع الملف ----------
uploaded = st.file_uploader("📤 اختار ملف التقرير", type=["pdf", "aspx"])

if uploaded is None:
    st.info("👆 لسه ما رفعتش ملف.")
    st.markdown(
        '<div class="app-footer">Powered by <b>Loans &amp; Economic Empowerment IT Group</b></div>',
        unsafe_allow_html=True,
    )
    st.stop()

# ---------- قراءة الملف ----------
try:
    with st.spinner("⏳ جاري قراءة الملف..."):
        result = core.parse_pdf(io.BytesIO(uploaded.getvalue()))
except Exception as e:
    st.error(f"⚠️ الملف ده مش تقرير سليم: {e}")
    st.stop()

client, loans, guarantors, installments, loan_list, unknown = result

# ---------- ملخص سريع ----------
c1, c2, c3, c4 = st.columns(4)
c1.metric("👤 العميل", client.get("الاسم", "—"))
c2.metric("🔢 الكود", client.get("الكود", "—"))
c3.metric("📅 عدد الأقساط", len([r for r in installments if r[1] != "إجمالي"]))
c4.metric("🤝 عدد الضامنين", len(guarantors))

if unknown:
    st.warning("⚠️ شكل جدول التمويلات في الملف ده مختلف شوية — اتحفظ في شيت \"غير معروف\" في الإكسيل.")

with st.expander("🔍 عرض بيانات العميل والتمويل"):
    st.write("**بيانات العميل**")
    st.table({"البيان": list(client.keys()), "القيمة": list(client.values())})
    if loans:
        st.write("**بيانات التمويل**")
        st.dataframe(loans, use_container_width=True)

st.divider()

tab_excel, tab_liq = st.tabs(["⬇️ تحويل لإكسيل", "🧮 مساعد التصفية"])

# ---------- تبويب الإكسيل ----------
with tab_excel:
    buf = io.BytesIO()
    core.write_workbook(result, buf)
    buf.seek(0)
    st.download_button(
        "⬇️ تنزيل ملف الإكسيل",
        data=buf,
        file_name=Path(uploaded.name).stem + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

# ---------- تبويب مساعد التصفية ----------
with tab_liq:
    groups = core.installments_by_loan(installments)
    if not groups:
        st.error("⚠️ مفيش جدول أقساط في التقرير ده.")
    else:
        codes = list(groups)
        code = codes[0] if len(codes) == 1 else st.selectbox("💳 اختار التمويل", codes)
        if not LIQ_HTML_PATH.exists():
            st.error(
                f"⚠️ مش لاقي صفحة مساعد التصفية ({LIQ_HTML_PATH.name}). "
                "تأكد إنها مرفوعة في نفس مستودع GitHub جنب app.py."
            )
        else:
            template = LIQ_HTML_PATH.read_text(encoding="utf-8")
            page_html = core.build_liquidation_html(
                template, groups[code], client, source_name=uploaded.name, loan_code=code
            )
            st.caption("✅ الأقساط جاهزة ومتعبّية — دوس الزرار يفتح مساعد التصفية في تبويب جديد.")
            open_in_new_tab_button("🧮 فتح مساعد التصفية في تبويب جديد", page_html, key=f"liq_{code}")
            st.info("💡 لو المتصفح منع فتح التبويب، دوس على أيقونة \"popup blocked\" جنب شريط العنوان واسمح له.")

# ---------- الفوتر ----------
st.markdown(
    '<div class="app-footer">Powered by <b>Loans &amp; Economic Empowerment IT Group</b></div>',
    unsafe_allow_html=True,
)
