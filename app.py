import streamlit as st
import pandas as pd
from gen import getProgress

# Please change normal SINGLE Backslash (\) to be DOUBLE Backslash (\\)
# for Path specify
folder = r"D:\\OneDrive - MPS\\BMC\\BMC-589\\03 PROJECT INFO\\094AZ1401-XX_MONTHLY REPORT\\data"
filename = r"Progress.xlsx"

xlsxfile = folder+"\\"+filename
progress = getProgress(xlsxfile=xlsxfile, sheetname="Sum-Calc")

df = pd.DataFrame(progress)
df["months"] = pd.to_datetime(df["months"], dayfirst=True)
df = df.sort_values(by="months")


@st.cache_data
def load_table(dataFrame):
    st.table(dataFrame)


@st.cache_data
def load_line(dataFrame):
    st.line_chart(dataFrame, x="months", y=["plan", "actual"])


st.write("""
# Project Status

This app is currently in development.

## Overall Work Progress
""")

load_table(df)
load_line(df)
