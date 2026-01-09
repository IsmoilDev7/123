import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Company Activity Dashboard", layout="wide")

st.title("📊 Company Activity Analysis Dashboard")
st.markdown("""
This dashboard provides a comprehensive analysis of company activity:
- Stage distribution
- Activity timeline and trends
- Source analysis
- Responsible person analysis
- Company distribution and address mapping
""")

# File uploader
file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])
if file:
    df = pd.read_excel(file)
    
    # Convert dates
    df['Date of creation'] = pd.to_datetime(df['Date of creation'], errors='coerce')
    df['Date modified'] = pd.to_datetime(df['Date modified'], errors='coerce')
    
    # Filters
    st.subheader("🔍 Filters")
    col1, col2, col3 = st.columns(3)
    with col1:
        stage_filter = st.multiselect("Stage", df['Stage'].unique(), default=df['Stage'].unique())
    with col2:
        source_filter = st.multiselect("Source", df['Source'].unique(), default=df['Source'].unique())
    with col3:
        responsible_filter = st.multiselect("Responsible", df['Responsible'].unique(), default=df['Responsible'].unique())
    
    df_filtered = df[df['Stage'].isin(stage_filter) & df['Source'].isin(source_filter) & df['Responsible'].isin(responsible_filter)]
    
    # KPI
    st.subheader("📌 Key Metrics")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Activities", len(df_filtered))
    k2.metric("Unique Companies", df_filtered['Company name'].nunique())
    k3.metric("Sources", df_filtered['Source'].nunique())
    k4.metric("Responsible Persons", df_filtered['Responsible'].nunique())
    
    # Stage Distribution
    st.subheader("📊 Stage Distribution")
    stage_count = df_filtered['Stage'].value_counts().reset_index()
    stage_count.columns = ['Stage','Count']
    st.plotly_chart(px.bar(stage_count, x='Stage', y='Count', color='Count',
                            title="Activities by Stage", color_continuous_scale='Blues'), use_container_width=True)
    
    # Activity Timeline
    st.subheader("📈 Activity Timeline")
    timeline = df_filtered.groupby('Date of creation').size().reset_index(name='Count')
    st.plotly_chart(px.line(timeline, x='Date of creation', y='Count', title="Activities over Time"), use_container_width=True)
    
    # Source Analysis
    st.subheader("🔎 Source Analysis")
    source_df = df_filtered['Source'].value_counts().reset_index()
    source_df.columns = ['Source','Count']
    st.plotly_chart(px.pie(source_df, names='Source', values='Count', title="Source Distribution"), use_container_width=True)
    
    # Responsible Person Analysis
    st.subheader("👤 Responsible Person Analysis")
    resp_df = df_filtered['Responsible'].value_counts().reset_index()
    resp_df.columns = ['Responsible','Count']
    st.plotly_chart(px.bar(resp_df, x='Responsible', y='Count', color='Count', title="Activities by Responsible Person", color_continuous_scale='Viridis'), use_container_width=True)
    
    # Company distribution
    st.subheader("🏢 Top Companies")
    company_df = df_filtered['Company name'].value_counts().reset_index().head(10)
    company_df.columns = ['Company','Count']
    st.plotly_chart(px.bar(company_df, x='Company', y='Count', color='Count', title="Top 10 Companies by Activities", color_continuous_scale='Plasma'), use_container_width=True)
    
    # Timeline by Stage
    st.subheader("📆 Timeline by Stage")
    timeline_stage = df_filtered.groupby(['Date of creation','Stage']).size().reset_index(name='Count')
    st.plotly_chart(px.area(timeline_stage, x='Date of creation', y='Count', color='Stage', title="Activities over Time by Stage"), use_container_width=True)
    
    # Comments word cloud (optional)
    try:
        from wordcloud import WordCloud
        import matplotlib.pyplot as plt
        st.subheader("💬 Comments Word Cloud")
        text = " ".join(str(c) for c in df_filtered['Comment'].dropna())
        wc = WordCloud(width=800, height=400, background_color="white").generate(text)
        fig, ax = plt.subplots(figsize=(10,5))
        ax.imshow(wc, interpolation='bilinear')
        ax.axis("off")
        st.pyplot(fig)
    except:
        st.info("WordCloud library not installed. Install wordcloud for comment visualization.")
    
    # Optional: Address mapping if geocoding available
    # st.subheader("📍 Company Address Mapping")
    # Use pydeck / plotly scatter_mapbox if lat/lon available
else:
    st.info("Upload your Excel file to generate all analyses automatically.")
