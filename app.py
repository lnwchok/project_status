import streamlit as st
import pandas as pd

# def load_progress():


st.write("""
# Project Status

This app is currently in development.

## Overall Work Progress
""")

df = pd.read_csv("progress.csv", names=["M","Plan","Actual"], index_col=None)
df["M"] = pd.to_datetime(df["M"], dayfirst=True)
df = df.sort_values(by="M")
st.table(df)
st.line_chart(df, x="M", y=["Plan", "Actual"])