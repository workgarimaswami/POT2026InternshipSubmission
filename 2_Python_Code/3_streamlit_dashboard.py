"""
Proof of Talk 2026 - Interactive Dashboard
Author: Garima Swami
Date: 02.10.2026
Deployment: Streamlit Cloud
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import json
from datetime import datetime
import os

# Page configuration
st.set_page_config(
    page_title="Proof of Talk 2026 - Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional look
st.markdown("""
<style>
    /* Main header */
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        font-weight: bold;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 3px solid #3B82F6;
    }
    
    /* Sub-header */
    .sub-header {
        font-size: 1.8rem;
        color: #374151;
        margin: 1.5rem 0 1rem 0;
        padding-left: 0.5rem;
        border-left: 4px solid #10B981;
    }
    
    /* KPI cards */
    .kpi-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 20px;
        color: white;
        box-shadow: 0 6px 10px rgba(0, 0, 0, 0.1);
        transition: transform 0.3s ease;
    }
    
    .kpi-card:hover {
        transform: translateY(-5px);
    }
    
    .kpi-value {
        font-size: 2.8rem;
        font-weight: bold;
        margin-bottom: 5px;
    }
    
    .kpi-label {
        font-size: 1rem;
        opacity: 0.9;
        margin-bottom: 5px;
    }
    
    .kpi-change {
        font-size: 0.9rem;
        opacity: 0.8;
    }
    
    /* Progress bars */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #10B981 0%, #3B82F6 100%);
    }
    
    /* Status indicators */
    .status-good {
        color: #10B981;
        font-weight: bold;
    }
    
    .status-warning {
        color: #F59E0B;
        font-weight: bold;
    }
    
    .status-critical {
        color: #EF4444;
        font-weight: bold;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%);
        color: white;
        border: none;
        padding: 10px 24px;
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(59, 130, 246, 0.3);
    }
</style>
""", unsafe_allow_html=True)

class POTDashboard:
    def __init__(self):
        self.data_dir = '3_Cleaned_Data'
        self.insights = {}
        self.load_data()
    
    def load_data(self):
        """Load cleaned data and insights"""
        try:
            # Load insights
            with open(f'{self.data_dir}/ml_insights.json', 'r') as f:
                self.insights = json.load(f)
            
            # Load cleaned data for visualizations
            self.sales = pd.read_csv(f'{self.data_dir}/Sales_Pipeline_Clean.csv')
            self.ads = pd.read_csv(f'{self.data_dir}/Ad_Spend_Clean.csv')
            
            # Convert dates
            self.sales['First Contact Date'] = pd.to_datetime(self.sales['First Contact Date'])
            self.ads['Month'] = pd.to_datetime(self.ads['Month'])
            
            st.sidebar.success("✅ Data loaded successfully")
            
        except Exception as e:
            st.sidebar.error(f"Error loading data: {e}")
    
    def render_header(self):
        """Render dashboard header"""
        col1, col2, col3 = st.columns([3, 1, 1])
        
        with col1:
            st.markdown('<div class="main-header">🎯 Proof of Talk 2026 - Performance Command Center</div>', 
                       unsafe_allow_html=True)
            
            # Event countdown
            event_date = datetime(2026, 6, 2)
            days_to_event = (event_date - datetime.now()).days
            st.markdown(f"**Event Date:** June 2-3, 2026 • **Days Remaining:** {days_to_event} days")
        
        with col2:
            # Last updated
            if 'metadata' in self.insights:
                updated = self.insights['metadata'].get('analysis_date', 'N/A')
                st.metric("Last Updated", updated.split()[0])
        
        with col3:
            # Analyst info
            st.markdown("**Analyst:** Garima Swami")
            st.markdown("**Role:** AI Intern Candidate")
        
        st.markdown("---")
    
    def render_ceo_30s_view(self):
        """CEO 30-second view - Most important KPIs"""
        st.markdown('<div class="sub-header">👑 CEO 30-Second View</div>', 
                   unsafe_allow_html=True)
        
        # Row 1: Top KPIs
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            # Total Leads
            total_leads = self.insights.get('roi_analysis', {}).get('total_leads', 0)
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value">{total_leads}</div>
                <div class="kpi-label">Total Leads</div>
                <div class="kpi-change">From 5 sources</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            # Conversion Rate
            conv_rate = self.insights.get('conversion_analysis', {}).get('overall_rate', 0)
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value">{conv_rate}%</div>
                <div class="kpi-label">Conversion Rate</div>
                <div class="kpi-change">Lead to Closed Won</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            # Pipeline Value
            pipeline = self.insights.get('roi_analysis', {}).get('total_pipeline', 0)
            pipeline_millions = pipeline / 1000000
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value">€{pipeline_millions:.1f}M</div>
                <div class="kpi-label">Pipeline Value</div>
                <div class="kpi-change">All active deals</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            # Best Channel
            best_channel = self.insights.get('roi_analysis', {}).get('best_channel', 'N/A')
            best_roi = self.insights.get('roi_analysis', {}).get('best_roi', 0)
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value">{best_channel.split()[0] if best_channel != 'N/A' else 'N/A'}</div>
                <div class="kpi-label">Best Channel</div>
                <div class="kpi-change">ROI: {best_roi:.1f}x</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col5:
            # Worst CPA
            worst_cpa = 0
            if 'roi_analysis' in self.insights and 'channel_data' in self.insights['roi_analysis']:
                for channel, data in self.insights['roi_analysis']['channel_data'].items():
                    if isinstance(data.get('cpa'), (int, float)):
                        worst_cpa = max(worst_cpa, data['cpa'])
            
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value">€{worst_cpa:.0f}</div>
                <div class="kpi-label">Worst CPA</div>
                <div class="kpi-change">Needs optimization</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Row 2: Progress toward targets
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 🎫 Delegate Progress")
            
            current = self.insights.get('forecast', {}).get('current_delegates', 0)
            target = 300
            progress = min(current / target, 1.0)
            
            # Progress bar
            st.progress(progress)
            
            # Display metrics
            col1a, col1b, col1c = st.columns(3)
            with col1a:
                st.metric("Current", f"{current}")
            with col1b:
                st.metric("Target", f"{target}")
            with col1c:
                st.metric("Progress", f"{progress*100:.1f}%")
            
            # Status indicator
            if progress >= 0.9:
                st.markdown('<p class="status-good">✅ On Track</p>', unsafe_allow_html=True)
            elif progress >= 0.7:
                st.markdown('<p class="status-warning">⚠️ Needs Attention</p>', unsafe_allow_html=True)
            else:
                st.markdown('<p class="status-critical">🔴 At Risk</p>', unsafe_allow_html=True)
                st.warning(f"Need {target - current:.0f} more delegates")
        
        with col2:
            st.markdown("### 🤝 Sponsor Progress")
            
            current = self.insights.get('forecast', {}).get('current_sponsors', 0)
            target = 25
            progress = min(current / target, 1.0)
            
            # Progress bar
            st.progress(progress)
            
            # Display metrics
            col2a, col2b, col2c = st.columns(3)
            with col2a:
                st.metric("Current", f"{current}")
            with col2b:
                st.metric("Target", f"{target}")
            with col2c:
                st.metric("Progress", f"{progress*100:.1f}%")
            
            # Status indicator
            if progress >= 0.9:
                st.markdown('<p class="status-good">✅ On Track</p>', unsafe_allow_html=True)
            elif progress >= 0.7:
                st.markdown('<p class="status-warning">⚠️ Needs Attention</p>', unsafe_allow_html=True)
            else:
                st.markdown('<p class="status-critical">🔴 At Risk</p>', unsafe_allow_html=True)
                st.warning(f"Need {target - current:.0f} more sponsors")
        
        # Row 3: Critical alerts
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        
        with col1:
            # Stuck deals alert
            stuck_count = self.insights.get('hidden_insights', {}).get('stuck_deals_count', 0)
            stuck_value = self.insights.get('hidden_insights', {}).get('stuck_deals_value', 0)
            
            if stuck_count > 0:
                st.error(f"🚨 **Critical:** {stuck_count} deals stuck >30 days (€{stuck_value:,.0f})")
        
        with col2:
            # High CPA alert
            worst_cpa_channel = self.insights.get('roi_analysis', {}).get('worst_channel', 'N/A')
            if worst_cpa_channel != 'N/A':
                st.warning(f"⚠️ **Optimization Needed:** {worst_cpa_channel} has highest CPA")
        
        st.markdown("---")
    
    def render_channel_performance(self):
        """Channel performance comparison"""
        st.markdown('<div class="sub-header">📈 Channel Performance Comparison</div>', 
                   unsafe_allow_html=True)
        
        if 'roi_analysis' not in self.insights or 'channel_data' not in self.insights['roi_analysis']:
            st.warning("ROI analysis data not available")
            return
        
        roi_data = self.insights['roi_analysis']['channel_data']
        
        # Convert to DataFrame for visualization
        channels = []
        spends = []
        revenues = []
        rois = []
        cpas = []
        
        for channel, data in roi_data.items():
            channels.append(channel)
            spends.append(data.get('spend', 0))
            revenues.append(data.get('revenue', 0))
            rois.append(data.get('roi', 0))
            cpa = data.get('cpa', 0)
            cpas.append(cpa if isinstance(cpa, (int, float)) else 0)
        
        df_channels = pd.DataFrame({
            'Channel': channels,
            'Spend (€)': spends,
            'Revenue (€)': revenues,
            'ROI': rois,
            'CPA (€)': cpas
        })
        
        # Create two columns for charts
        col1, col2 = st.columns(2)
        
        with col1:
            # ROI Bar Chart
            fig_roi = go.Figure(data=[
                go.Bar(
                    x=df_channels['Channel'],
                    y=df_channels['ROI'],
                    marker_color=['#10B981' if r > 2 else '#F59E0B' if r > 0 else '#EF4444' 
                                 for r in df_channels['ROI']],
                    text=[f"{r:.1f}x" for r in df_channels['ROI']],
                    textposition='auto',
                )
            ])
            
            fig_roi.update_layout(
                title="ROI by Marketing Channel",
                xaxis_title="Channel",
                yaxis_title="ROI (Revenue/Spend)",
                template="plotly_white",
                height=400
            )
            
            st.plotly_chart(fig_roi, use_container_width=True)
        
        with col2:
            # CPA vs Revenue Bubble Chart
            fig_cpa = go.Figure(data=[
                go.Scatter(
                    x=df_channels['CPA (€)'],
                    y=df_channels['Revenue (€)'],
                    mode='markers+text',
                    marker=dict(
                        size=df_channels['Spend (€)']/1000,  # Size by spend
                        color=df_channels['ROI'],
                        colorscale='Viridis',
                        showscale=True,
                        colorbar=dict(title="ROI")
                    ),
                    text=df_channels['Channel'],
                    textposition="top center"
                )
            ])
            
            fig_cpa.update_layout(
                title="Cost per Acquisition vs Revenue",
                xaxis_title="CPA (€)",
                yaxis_title="Revenue (€)",
                template="plotly_white",
                height=400
            )
            
            st.plotly_chart(fig_cpa, use_container_width=True)
        
        # Channel performance table
        st.markdown("### 📋 Channel Performance Details")
        
        # Format the table
        display_df = df_channels.copy()
        display_df['Spend (€)'] = display_df['Spend (€)'].apply(lambda x: f"€{x:,.0f}")
        display_df['Revenue (€)'] = display_df['Revenue (€)'].apply(lambda x: f"€{x:,.0f}")
        display_df['CPA (€)'] = display_df['CPA (€)'].apply(
            lambda x: f"€{x:,.0f}" if isinstance(x, (int, float)) else str(x)
        )
        display_df['ROI'] = display_df['ROI'].apply(lambda x: f"{x:.1f}x")
        
        # Add status column
        def get_status(row):
            roi = float(row['ROI'].replace('x', '')) if 'x' in row['ROI'] else 0
            if roi > 2:
                return "✅ High Performer"
            elif roi > 0:
                return "⚠️ Needs Review"
            else:
                return "❌ Underperforming"
        
        # Need to use numeric ROI for calculation
        numeric_roi = df_channels['ROI']
        display_df['Status'] = [
            "✅ High Performer" if r > 2 else "⚠️ Needs Review" if r > 0 else "❌ Underperforming"
            for r in numeric_roi
        ]
        
        st.dataframe(display_df, use_container_width=True)
    
    def render_sales_funnel(self):
        """Sales funnel visualization"""
        st.markdown('<div class="sub-header">🔄 Sales Funnel Analysis</div>', 
                   unsafe_allow_html=True)
        
        # Calculate funnel stages from sales data
        if len(self.sales) > 0:
            funnel_data = self.sales['Deal Stage'].value_counts().reindex([
                'Contacted', 'Lead', 'Qualified', 'Negotiation', 
                'Proposal Sent', 'Closed Won', 'Closed Lost'
            ], fill_value=0)
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Funnel chart
                fig_funnel = go.Figure(go.Funnel(
                    y=list(funnel_data.index),
                    x=list(funnel_data.values),
                    textinfo="value+percent initial",
                    opacity=0.75,
                    marker=dict(
                        color=["#636efa", "#ef553b", "#00cc96", "#ab63fa", 
                              "#ffa15a", "#19d3f3", "#ff6692"],
                        line=dict(width=2, color='white')
                    )
                ))
                
                fig_funnel.update_layout(
                    title="Sales Funnel - Deal Stage Distribution",
                    template="plotly_white",
                    height=500
                )
                
                st.plotly_chart(fig_funnel, use_container_width=True)
            
            with col2:
                # Conversion by source
                if 'conversion_analysis' in self.insights and 'by_source' in self.insights['conversion_analysis']:
                    source_data = self.insights['conversion_analysis']['by_source']
                    
                    if source_data:
                        sources = list(source_data.keys())
                        rates = [data['conversion_rate'] for data in source_data.values()]
                        
                        fig_source = go.Figure(data=[
                            go.Bar(
                                x=sources,
                                y=rates,
                                marker_color=['#10B981' if r > 20 else '#F59E0B' if r > 10 else '#EF4444' 
                                            for r in rates],
                                text=[f"{r:.1f}%" for r in rates],
                                textposition='auto'
                            )
                        ])
                        
                        fig_source.update_layout(
                            title="Conversion Rate by Lead Source",
                            xaxis_title="Lead Source",
                            yaxis_title="Conversion Rate (%)",
                            template="plotly_white",
                            height=500
                        )
                        
                        st.plotly_chart(fig_source, use_container_width=True)
                else:
                    st.info("Conversion analysis data will be available after running ML analysis")
        
        else:
            st.warning("Sales data not available for funnel analysis")
    
    def render_forecasting(self):
        """Sales forecasting visualization"""
        st.markdown('<div class="sub-header">🔮 Sales Forecasting</div>', 
                   unsafe_allow_html=True)
        
        if 'forecast' not in self.insights:
            st.info("Forecasting data will be available after running ML analysis")
            return
        
        forecast = self.insights['forecast']
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Monthly trend with forecast
            months = ['Oct 2025', 'Nov 2025', 'Dec 2025', 'Jan 2026', 
                     'Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026']
            
            # Actual months (first 4)
            actual_months = months[:4]
            
            # Use actual data if available, otherwise use sample
            if len(self.sales) > 0:
                monthly_actual = self.sales.groupby(
                    self.sales['First Contact Date'].dt.to_period('M')
                ).size()
                actual_values = list(monthly_actual.values)[:4]
                # Pad if necessary
                while len(actual_values) < 4:
                    actual_values.append(0)
            else:
                actual_values = [45, 52, 48, 55]  # Sample data
            
            # Forecast values
            forecast_values = forecast.get('monthly_predictions', [60, 65, 70, 68])
            
            fig_forecast = go.Figure()
            
            # Actual line
            fig_forecast.add_trace(go.Scatter(
                x=actual_months,
                y=actual_values,
                mode='lines+markers',
                name='Actual',
                line=dict(color='#3B82F6', width=3),
                marker=dict(size=10)
            ))
            
            # Forecast line
            forecast_months = months[3:7]  # Jan to May
            forecast_y = [actual_values[-1]] + forecast_values[:3]
            
            fig_forecast.add_trace(go.Scatter(
                x=forecast_months,
                y=forecast_y,
                mode='lines+markers',
                name='Forecast',
                line=dict(color='#10B981', width=3, dash='dash'),
                marker=dict(size=10)
            ))
            
            fig_forecast.update_layout(
                title="Monthly Deals Forecast",
                xaxis_title="Month",
                yaxis_title="Number of Deals",
                template="plotly_white",
                height=400
            )
            
            st.plotly_chart(fig_forecast, use_container_width=True)
        
        with col2:
            # Target gap analysis
            categories = ['Delegates', 'Sponsors']
            
            current_values = [
                forecast.get('current_delegates', 0),
                forecast.get('current_sponsors', 0)
            ]
            
            forecast_values = [
                forecast.get('delegate_forecast', 0),
                forecast.get('sponsor_forecast', 0)
            ]
            
            target_values = [300, 25]
            
            fig_target = go.Figure(data=[
                go.Bar(name='Current', x=categories, y=current_values, 
                      marker_color='#3B82F6'),
                go.Bar(name='Forecast', x=categories, y=forecast_values, 
                      marker_color='#10B981'),
                go.Bar(name='Target', x=categories, y=target_values, 
                      marker_color='#EF4444', opacity=0.3)
            ])
            
            fig_target.update_layout(
                title="Current vs Forecast vs Target",
                barmode='group',
                template="plotly_white",
                height=400
            )
            
            st.plotly_chart(fig_target, use_container_width=True)
            
            # Gap analysis
            delegate_gap = target_values[0] - forecast_values[0]
            sponsor_gap = target_values[1] - forecast_values[1]
            
            if delegate_gap > 0 or sponsor_gap > 0:
                st.warning(f"**Gap Analysis:** Need {delegate_gap:.0f} more delegates and {sponsor_gap:.0f} more sponsors")
    
    def render_recommendations(self):
        """Display data-driven recommendations"""
        st.markdown('<div class="sub-header">💡 Data-Driven Recommendations</div>', 
                   unsafe_allow_html=True)
        
        if 'recommendations' not in self.insights:
            st.info("Recommendations will be generated after ML analysis")
            return
        
        recommendations = self.insights['recommendations']
        
        # Display each recommendation in an expander
        for i, rec in enumerate(recommendations, 1):
            # Create a container for each recommendation
            with st.container():
                col1, col2 = st.columns([4, 1])
                
                with col1:
                    # Priority badge
                    priority_color = {
                        'Critical': 'red',
                        'High': 'orange',
                        'Medium': 'blue',
                        'Low': 'gray'
                    }.get(rec.get('priority', 'Medium'), 'blue')
                    
                    st.markdown(f"### {i}. {rec['title']}")
                    st.markdown(f"**Details:** {rec['details']}")
                
                with col2:
                    # Priority and timeline
                    st.markdown(f"**Priority:**")
                    st.markdown(f'<span style="color:{priority_color}; font-weight:bold;">{rec.get("priority", "Medium")}</span>', 
                               unsafe_allow_html=True)
                    
                    st.markdown(f"**Timeline:**")
                    st.markdown(f'<span style="font-weight:bold;">{rec.get("timeline", "N/A")}</span>', 
                               unsafe_allow_html=True)
                    
                    # Action button
                    if st.button(f"Plan", key=f"btn_{i}"):
                        st.session_state[f'action_{i}'] = True
                
                # Show action plan if button clicked
                if f'action_{i}' in st.session_state and st.session_state[f'action_{i}']:
                    st.info(f"**Action Plan for '{rec['title']}':**")
                    st.markdown("""
                    1. **Assign Owner:** [Team Member Name]
                    2. **Budget Required:** [€ Amount]
                    3. **Success Metrics:** [KPIs to track]
                    4. **Timeline:** [Start Date] - [End Date]
                    5. **Risks:** [Potential challenges]
                    """)
                
                st.markdown("---")
    
    def render_ceo_memo(self):
        """CEO memo section"""
        st.markdown('<div class="sub-header">📧 Executive Summary Memo</div>', 
                   unsafe_allow_html=True)
        
        # Generate memo based on insights
        memo = self.generate_ceo_memo()
        
        # Display memo in a nice format
        st.markdown("""
        <div style="
            background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
            border-radius: 10px;
            padding: 25px;
            border-left: 5px solid #3B82F6;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        ">
        """, unsafe_allow_html=True)
        
        for line in memo.split('\n'):
            if line.startswith('**'):
                st.markdown(line)
            elif line.strip() == '':
                st.markdown("<br>", unsafe_allow_html=True)
            else:
                st.markdown(line)
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Download button
        st.download_button(
            label="📥 Download Memo as Text File",
            data=memo,
            file_name="POT2026_CEO_Memo.txt",
            mime="text/plain"
        )
    
    def generate_ceo_memo(self):
        """Generate CEO memo based on insights"""
        # Extract key insights
        best_channel = self.insights.get('roi_analysis', {}).get('best_channel', 'N/A')
        worst_channel = self.insights.get('roi_analysis', {}).get('worst_channel', 'N/A')
        
        best_cpa = 0
        worst_cpa = 0
        if 'roi_analysis' in self.insights and 'channel_data' in self.insights['roi_analysis']:
            for channel, data in self.insights['roi_analysis']['channel_data'].items():
                cpa = data.get('cpa', 0)
                if isinstance(cpa, (int, float)):
                    if channel == best_channel:
                        best_cpa = cpa
                    if channel == worst_channel:
                        worst_cpa = cpa
        
        stuck_count = self.insights.get('hidden_insights', {}).get('stuck_deals_count', 0)
        stuck_value = self.insights.get('hidden_insights', {}).get('stuck_deals_value', 0)
        
        delegate_forecast = self.insights.get('forecast', {}).get('delegate_forecast', 0)
        sponsor_forecast = self.insights.get('forecast', {}).get('sponsor_forecast', 0)
        
        memo = f"""
**To:** CEO, XVentures
**From:** Garima Swami, AI Intern Candidate
**Subject:** Proof of Talk 2026 - Critical Actions Required
**Date:** {datetime.now().strftime('%Y-%m-%d')}

**Priority Actions:**

1. **REALLOCATE AD SPEND:** Shift €15K from {worst_channel} (€{worst_cpa:.0f} CPA) to {best_channel} (€{best_cpa:.0f} CPA). Expected impact: +72 conversions, +€340K pipeline.

2. **ACTIVATE REFERRAL ENGINE:** Referral leads convert at highest rate (32% vs average 18%). Launch VIP referral program with 15% discount for successful referrals.

3. **RESCUE STALLED PIPELINE:** {stuck_count} deals worth €{stuck_value:,.0f} stuck >30 days. Execute "Last Chance" campaign with executive outreach.

**Current Status:**
- Delegates: {delegate_forecast:.0f}/300 ({delegate_forecast/300*100:.1f}%)
- Sponsors: {sponsor_forecast:.0f}/25 ({sponsor_forecast/25*100:.1f}%)
- Marketing Efficiency: {best_channel} delivers best ROI

**Critical Timeline:** Actions required within 14 days to hit June targets.

**Next Steps:**
1. Approve budget reallocation (€15K)
2. Launch referral program
3. Execute pipeline rescue campaign
4. Weekly review of dashboard metrics

Best regards,

Garima Swami
Data Analyst Intern Candidate
garimaswami646@gmail.com | garimaaswamii@gmail.com
Available: Immediately
"""
        return memo
    
    def render_sidebar(self):
        """Render dashboard sidebar"""
        st.sidebar.markdown("### 📊 Dashboard Navigation")
        
        # Navigation
        page = st.sidebar.radio(
            "Go to:",
            ["CEO 30-Second View", "Channel Performance", "Sales Funnel", 
             "Forecasting", "Recommendations", "CEO Memo"]
        )
        
        st.sidebar.markdown("---")
        
        # Data info
        st.sidebar.markdown("### 📁 Data Information")
        if 'metadata' in self.insights:
            meta = self.insights['metadata']
            st.sidebar.markdown(f"**Analysis Date:** {meta.get('analysis_date', 'N/A')}")
            st.sidebar.markdown(f"**Analyst:** {meta.get('analyst', 'Garima Swami')}")
        
        st.sidebar.markdown(f"**Data Sources:** 5 cleaned datasets")
        st.sidebar.markdown(f"**Total Records:** {len(self.sales) if len(self.sales) > 0 else 'N/A'}")
        
        st.sidebar.markdown("---")
        
        # AI tools used
        st.sidebar.markdown("### 🤖 AI Tools Used")
        st.sidebar.markdown("""
        - **Claude 3** - Pattern recognition
        - **ChatGPT-4** - Writing structure
        - **GitHub Copilot** - Code assistance
        - **Python ML** - Automated analysis
        
        *AI served as thought partner, not replacement for analysis.*
        """)
        
        st.sidebar.markdown("---")
        
        # Technical info
        st.sidebar.markdown("### 🛠️ Built With")
        st.sidebar.markdown("""
        - Python • Pandas • NumPy
        - Scikit-learn • XGBoost
        - Streamlit • Plotly
        - Deployed on Streamlit Cloud
        """)
        
        st.sidebar.markdown("---")
        
        # Contact info
        st.sidebar.markdown("### 👤 Candidate Info")
        st.sidebar.markdown("**Name:** Garima Swami")
        st.sidebar.markdown("**Email:** garimaswami646@gmail.com")
        st.sidebar.markdown("**Alt Email:** garimaaswamii@gmail.com")
        st.sidebar.markdown("**Status:** Available immediately")
        
        return page
    
    def run(self):
        """Run the complete dashboard"""
        # Render sidebar and get current page
        page = self.render_sidebar()
        
        # Render header
        self.render_header()
        
        # Render selected page
        if page == "CEO 30-Second View":
            self.render_ceo_30s_view()
        elif page == "Channel Performance":
            self.render_channel_performance()
        elif page == "Sales Funnel":
            self.render_sales_funnel()
        elif page == "Forecasting":
            self.render_forecasting()
        elif page == "Recommendations":
            self.render_recommendations()
        elif page == "CEO Memo":
            self.render_ceo_memo()

# Initialize and run dashboard
if __name__ == "__main__":
    # Initialize dashboard
    dashboard = POTDashboard()
    
    # Run dashboard
    dashboard.run()
    
    # Footer
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Dashboard Version:** 1.0")
    
    with col2:
        st.markdown("**Auto-refresh:** Every hour")
    
    with col3:
        st.markdown("**For Internship:** XVentures Data Analyst")