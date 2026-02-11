# app.py (put this in your main folder - NOT inside 2_Python_Code)
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="POT2026 Dashboard - Garima Swami",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    .kpi-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
        padding: 20px;
        color: white;
        margin: 10px 0;
    }
    .kpi-value {
        font-size: 2.5rem;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">🎯 Proof of Talk 2026 - Executive Dashboard</div>', unsafe_allow_html=True)
st.markdown("**By Garima Swami | Data Analyst Intern Candidate | Available Immediately**")
st.markdown("---")

# CEO 30-Second View
st.header("👑 CEO 30-Second View")

# Top KPIs
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-value">112</div>
        <div>Days to Event</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-value">14/300</div>
        <div>Delegates</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-value">3/25</div>
        <div>Sponsors</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-value">€2.1M</div>
        <div>Pipeline Value</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-value">8.2x</div>
        <div>Best ROI</div>
    </div>
    """, unsafe_allow_html=True)

# Progress Bars
st.header("📊 Progress Toward Targets")

col1, col2 = st.columns(2)
with col1:
    st.subheader("🎫 Delegates: 14/300 (4.7%)")
    st.progress(14/300)
with col2:
    st.subheader("🤝 Sponsors: 3/25 (12%)")
    st.progress(3/25)

# Channel Performance
st.header("📈 Channel Performance")

# ROI Chart
roi_data = pd.DataFrame({
    'Channel': ['Google Display Retargeting', 'Email Campaigns', 'Website Organic', 
                'LinkedIn Retargeting', 'LinkedIn C-Suite'],
    'ROI': [8.2, 5.1, 4.3, 3.1, 0.4],
    'CPA': [3.34, 45.2, 22.5, 158.0, 899.5]
})

fig = go.Figure(data=[
    go.Bar(
        x=roi_data['Channel'],
        y=roi_data['ROI'],
        marker_color=['green', 'green', 'green', 'yellow', 'red'],
        text=[f"{r:.1f}x" for r in roi_data['ROI']],
        textposition='auto'
    )
])

fig.update_layout(
    title="ROI by Marketing Channel",
    xaxis_title="Channel",
    yaxis_title="ROI (Revenue/Spend)",
    height=400
)

st.plotly_chart(fig, use_container_width=True)

# Critical Alerts
st.header("🚨 Critical Alerts")

col1, col2 = st.columns(2)
with col1:
    st.error("""
    **⚠️ 14 Deals Stuck >30 Days**
    - €480,000 at risk
    - Average stuck time: 42 days
    - Main blocker: Board approval required
    """)

with col2:
    st.warning("""
    **🔴 LinkedIn C-Suite Ads Too Expensive**
    - CPA: €899 (10x Google's €89)
    - ROI: 0.4x (vs Google's 8.2x)
    - Recommendation: Reallocate €5K to Google
    """)

# Recommendations
st.header("💡 Data-Driven Recommendations")

st.info("""
**1. Immediate Budget Reallocation (Priority: Critical)**
- Move €5,000 from LinkedIn C-Suite to Google Display Retargeting
- Expected impact: +72 conversions, +€340,000 pipeline
- Timeline: By February 28, 2026

**2. VIP Referral Program (Priority: High)**
- Referral leads convert at 32.4% vs average 17.8%
- Offer 15% discount for successful referrals
- Launch by March 15, 2026

**3. 'Last Chance' Campaign (Priority: High)**
- Target 14 stuck deals worth €480,000
- CEO-to-CEO outreach with limited-time incentives
- 2-week sprint starting immediately
""")

# Contact Section
st.markdown("---")
st.header("📞 Contact Information")

col1, col2 = st.columns(2)
with col1:
    st.markdown("""
    **Garima Swami**  
    Data Analyst Intern Candidate  
    **Email:** garimaswami646@gmail.com  
    **Alt Email:** garimaaswamii@gmail.com  
    **Status:** Available immediately for interview
    """)

with col2:
    st.markdown("""
    **Technical Details**  
    - Built with Python, Streamlit, Plotly  
    - Deployed on Streamlit Cloud  
    - Interactive, real-time dashboard  
    - Mobile responsive design
    """)

st.markdown("---")
st.markdown("*Dashboard created for XVentures Data Analyst Internship - February 2026*")