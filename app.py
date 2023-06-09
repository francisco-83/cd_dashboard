import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Customer Support Dashboard",
    page_icon = ":bar_chart:",
    layout = "wide"

)

@st.cache_data
def get_data_from_excel():
    df = pd.read_excel(
        io ="call_log_01.xlsx",
        engine='openpyxl',
        skiprows='1',
        usecols='A:L',
        nrows=2434,
    )
    return df
df = get_data_from_excel()

st.sidebar.header("Please filter here")

date = st.sidebar.multiselect(
    "Select the Date:",
    options=df["Month"].unique(),
    default=df["Month"].unique(),
)

topic = st.sidebar.multiselect(
    "Select the Topic:",
    options=df["Topic"].unique(),
    default=df["Topic"].unique()

)

operator = st.sidebar.multiselect(
    "Select the Operator:",
    options=df["Operator"].unique(),
    default=df["Operator"].unique()

)

resolved = st.sidebar.multiselect(
    "Select if Resolved:",
    options=df["Resolved"].unique(),
    default=df["Resolved"].unique()

)

df_selection = df.query(
    "Topic == @topic & Operator == @operator & Resolved == @resolved & Month == @date"
)

# Title
st.title(":bar_chart: Customer Support Dashboard")
st.markdown('##')

#Columns

left_column, center_column, right_column  = st.columns(3)

# Total Columns
total_calls = len(df_selection)
with left_column:
    st.subheader(":telephone_receiver: Total calls")
    st.subheader(total_calls)

#Resolved Calls
resolved = pd.value_counts(df_selection['Resolved'])
with center_column:
    st.subheader("✅ Resolved Calls")
    st.subheader(resolved[0])

# Unresolved calls
unresolved = pd.value_counts(df_selection['Resolved'])
with right_column:
    st.subheader("❌ Unresolved Calls")
    st.subheader(resolved[1])

st.markdown('---')

# Calls by operator

column_one, column_two = st.columns(2)

operator = pd.value_counts(df_selection['Operator'])

with column_one:
    fig_operator = px.bar(
        operator,
        x=operator,
        y=operator.index,
        orientation="h",
        width=500, 
        height=400,
        title="<b>Num. of Calls By Operator</b>",
        template="plotly"
    )
    st.plotly_chart(fig_operator)

topics = pd.value_counts(df_selection["Topic"])

with column_two:
    fig_topics = px.bar(
        topics,
        x=topics,
        y=topics.index,
        width=600, 
        height=400,
        orientation="h",
        title="<b>Num. of Calls By Topic</b>",
        template="plotly"
    )
    st.plotly_chart(fig_topics)
