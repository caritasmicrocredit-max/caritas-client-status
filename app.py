# -*- coding: utf-8 -*-
"""
مساعد حالة العميل — نسخة ويب (Streamlit)
==========================================
نفس فكرة البرنامج المكتبي، بس شغّال من المتصفح عشان أي حد في الفريق (في أي محافظة) يستخدمه
من غير ما يثبّت بايثون: يرفع ملف التقرير (PDF أو ASPX)، وياخد إكسيل جاهز، أو يفتح مساعد
التصفية وهو متعبّي بالأقساط تلقائي.

التشغيل محليًا:  streamlit run app.py
الملفات المطلوبة جنب app.py: core.py + مساعد_تصفية_العملاء_من_برنامج_المحصل.html
"""
import io
from pathlib import Path

import streamlit as st

import core

st.set_page_config(page_title="مساعد حالة العميل", page_icon="📄", layout="centered")

# اتجاه الصفحة من اليمين لليسار
st.markdown(
    "<style>html, body, [class*='css'] { direction: rtl; text-align: right; }</style>",
    unsafe_allow_html=True,
)

LIQ_HTML_PATH = Path(__file__).parent / "مساعد_تصفية_العملاء_من_برنامج_المحصل.html"

st.title("📄 مساعد حالة العميل")
st.caption("ارفع تقرير \"حالة العميل\" (PDF أو ASPX) وحوّله لإكسيل، أو افتح مساعد التصفية بالأقساط جاهزة.")

uploaded = st.file_uploader("اختار ملف التقرير", type=["pdf", "aspx"])

if uploaded is None:
    st.info("لسه ما رفعتش ملف.")
    st.stop()

# ---------- قراءة الملف ----------
try:
    with st.spinner("جاري قراءة الملف..."):
        result = core.parse_pdf(io.BytesIO(uploaded.getvalue()))
except Exception as e:
    st.error(f"الملف ده مش تقرير سليم: {e}")
    st.stop()

client, loans, guarantors, installments, loan_list, unknown = result

# ---------- ملخص سريع ----------
c1, c2, c3, c4 = st.columns(4)
c1.metric("العميل", client.get("الاسم", "—"))
c2.metric("الكود", client.get("الكود", "—"))
c3.metric("عدد الأقساط", len([r for r in installments if r[1] != "إجمالي"]))
c4.metric("عدد الضامنين", len(guarantors))

if unknown:
    st.warning("شكل جدول القروض في الملف ده مختلف شوية — اتحفظ في شيت \"غير معروف\" في الإكسيل.")

with st.expander("عرض بيانات العميل والقرض"):
    st.write("**بيانات العميل**")
    st.table({"البيان": list(client.keys()), "القيمة": list(client.values())})
    if loans:
        st.write("**بيانات القرض**")
        st.dataframe(loans, use_container_width=True)

st.divider()

tab_excel, tab_liq = st.tabs(["⬇️ تحويل لإكسيل", "🧮 مساعد التصفية"])

# ---------- تبويب الإكسيل ----------
with tab_excel:
    buf = io.BytesIO()
    core.write_workbook(result, buf)
    buf.seek(0)
    st.download_button(
        "تنزيل ملف الإكسيل",
        data=buf,
        file_name=Path(uploaded.name).stem + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

# ---------- تبويب مساعد التصفية ----------
with tab_liq:
    groups = core.installments_by_loan(installments)
    if not groups:
        st.error("مفيش جدول أقساط في التقرير ده.")
    else:
        codes = list(groups)
        code = codes[0] if len(codes) == 1 else st.selectbox("اختار القرض", codes)
        if not LIQ_HTML_PATH.exists():
            st.error(
                f"مش لاقي صفحة مساعد التصفية ({LIQ_HTML_PATH.name}). "
                "تأكد إنها مرفوعة في نفس مستودع GitHub جنب app.py."
            )
        else:
            template = LIQ_HTML_PATH.read_text(encoding="utf-8")
            page_html = core.build_liquidation_html(
                template, groups[code], client, source_name=uploaded.name, loan_code=code
            )
            st.components.v1.html(page_html, height=1400, scrolling=True)
