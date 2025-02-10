import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Please change normal SINGLE Backslash (\) to be DOUBLE Backslash (\\)
# for Path specify

xlsxfile = r"D:\\OneDrive - MPS\\BMC\\BMC-589\\00 APP\\project_status\\data.xlsx"

df = pd.read_excel(
    xlsxfile, sheet_name="progress", header=None, names=["months", "plan", "actual"]
)
df["months"] = df["months"].dt.strftime("%b %y")
df["plan"] = df["plan"].map("{:.1%}".format)
df["actual"] = df["actual"].map("{:.1%}".format)

# NOTE:No Longer Use
#
# @st.cache_data
# def load_line():
#     # st.line_chart(df, x="months", y=["plan", "actual"])
#     fig = px.line(
#         df,
#         x="months",
#         y=["plan", "actual"],
#         title="Overall Progress",
#         markers=True,
#         color_discrete_map={"plan": "#b5b5b5", "actual": "#ff2b2b"},
#     )
#     fig.update_layout(legend_title="Status")

#     st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})


@st.cache_data
def load_line_scatter(df):

    fig = go.Figure()

    # Plan Value
    fig.add_trace(
        go.Scatter(
            x=df["months"],
            y=df["plan"],
            mode="lines+markers+text",
            name="Plan",
            line=dict(color="#b5b5b5", width=1, dash="dash"),
            text=df["plan"],
            textposition="bottom right",
            textfont=dict(color="#b5b5b5")
        )
    )
    # Actual Value
    fig.add_trace(
        go.Scatter(
            x=df["months"],
            y=df["actual"],
            mode="lines+markers+text",
            name="Actual",
            line=dict(color="red", width=4),
            text=df["actual"],
            textposition="top left",
            textfont=dict(color="red")
        )
    )
    fig.update_layout(
        title="Overall Progress",  # Title
        yaxis_title="%-Progress",  # x-axis name
        xaxis_tickangle=-45,  # Set the x-axis label angle
        showlegend=True,  # Display the legend
        # plot_bgcolor="white",  # Set the plot background color
        # Set the paper (outside plot area) background color
        # paper_bgcolor="lightblue",
    )
    st.plotly_chart(fig, use_container_width=True, config={"staticPlot": True})



st.write(
    """
# Project Status

This app is currently in development.
"""
)

tab1, tab2 = st.tabs(["Progress", "Resources"])

with tab1:
    st.write("## Overall Work Progress")
    with st.expander("Data Preview"):
        st.dataframe(df, use_container_width=True)

    load_line_scatter(df=df)