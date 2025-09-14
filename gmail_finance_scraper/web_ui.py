import streamlit as st
from .scanner import run_scraper

st.title("Gmail Finance Scraper")

with st.form("options"):
    days = st.number_input("Look back N days", min_value=0, value=30)
    exclude = st.text_input("Exclude labels (comma separated)", "")
    min_amount = st.number_input("Minimum amount", min_value=0.0, value=0.0)
    submitted = st.form_submit_button("Scan")

if submitted:
    labels = [l.strip() for l in exclude.split(",") if l.strip()]
    df = run_scraper(days=days, exclude_labels=labels, min_amount=min_amount)
    st.dataframe(df)
    if st.button("Export to Excel"):
        df.to_excel("email_amounts.xlsx", index=False)
        st.success("Saved to email_amounts.xlsx")
