"""
Proof of Talk 2026 - Data Cleaning Pipeline
Author: Garima Swami
Date: 02.10.2026
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')
import os

class POTDataCleaner:
    def __init__(self, excel_path='POT2026_Raw_Data_Case_Study.xlsx'):
        self.excel_path = excel_path
        self.data = {}
        self.cleaning_log = []
        
    def log_step(self, message):
        """Log cleaning steps"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Remove emojis from log message
        import re
        message_no_emoji = re.sub(r'[^\x00-\x7F]+', '', message)
        self.cleaning_log.append(f"[{timestamp}] {message_no_emoji}")
        print(message)  # Keep emojis in console output
        
    def load_all_sheets(self):
        """Load all 5 sheets from Excel"""
        self.log_step("📂 Loading all data sheets...")
        
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(self.excel_path)
            sheet_names = excel_file.sheet_names
            
            for sheet in sheet_names:
                self.data[sheet] = pd.read_excel(self.excel_path, sheet_name=sheet)
                self.log_step(f"   ✓ Loaded: {sheet} ({len(self.data[sheet])} rows)")
                
        except Exception as e:
            self.log_step(f"   ✗ Error loading Excel: {e}")
            raise
        
        return self
    
    def clean_website_traffic(self):
        """Clean Website Traffic sheet"""
        self.log_step("🧹 Cleaning Website Traffic data...")
        df = self.data['Website Traffic'].copy()
        
        # 1. Remove duplicates
        initial_rows = len(df)
        df = df.drop_duplicates()
        removed_duplicates = initial_rows - len(df)
        if removed_duplicates > 0:
            self.log_step(f"   Removed {removed_duplicates} duplicate rows")
        
        # 2. Fix bounce rate outliers (values > 1 should be decimals)
        bounce_rate_fixes = 0
        for idx, row in df.iterrows():
            bounce = row['Bounce Rate']
            if isinstance(bounce, (int, float)):
                if bounce > 1:
                    df.at[idx, 'Bounce Rate'] = bounce / 100
                    bounce_rate_fixes += 1
        
        if bounce_rate_fixes > 0:
            self.log_step(f"   Fixed {bounce_rate_fixes} bounce rate values (>1)")
        
        # 3. Calculate missing conversion rates
        missing_conversions = df['Conversion Rate'].isna().sum()
        if missing_conversions > 0:
            mask = df['Conversion Rate'].isna()
            df.loc[mask, 'Conversion Rate'] = (
                df.loc[mask, 'Ticket Inquiry Conversions'] / 
                df.loc[mask, 'Sessions'] * 100
            )
            self.log_step(f"   Calculated {missing_conversions} missing conversion rates")
        
        # 4. Standardize traffic source names
        df['Traffic Source'] = df['Traffic Source'].str.lower().str.replace(' ', '_')
        
        # 5. Ensure correct data types
        df['Week Starting'] = pd.to_datetime(df['Week Starting'])
        numeric_cols = ['Sessions', 'Users', 'New Users', 'Ticket Inquiry Conversions']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        self.data['Website Traffic Clean'] = df
        self.log_step("   ✓ Website Traffic cleaned successfully")
        return self
    
    def clean_social_media(self):
        """Clean Social Media sheet"""
        self.log_step("📱 Cleaning Social Media data...")
        df = self.data['Social Media'].copy()
        
        # 1. Standardize platform names
        platform_mapping = {
            'X/Twitter': 'Twitter',
            'Twitter': 'Twitter',
            'X': 'Twitter'
        }
        df['Platform'] = df['Platform'].replace(platform_mapping)
        
        # 2. Fix engagement rates (handle percentages and decimals)
        def fix_engagement_rate(value):
            if pd.isna(value):
                return np.nan
            if isinstance(value, str):
                # Remove % sign and convert
                value = value.replace('%', '').strip()
                try:
                    num_value = float(value)
                    # If it looks like a percentage (0-100), convert to decimal
                    if 0 <= num_value <= 100:
                        return num_value / 100
                    return num_value
                except:
                    return np.nan
            elif isinstance(value, (int, float)):
                # If value > 1, assume it's percentage
                if value > 1:
                    return value / 100
                return value
            return np.nan
        
        df['Engagement Rate'] = df['Engagement Rate'].apply(fix_engagement_rate)
        
        # 3. Handle missing/NA values
        df['Top Post Impressions'] = df['Top Post Impressions'].replace(
            ['N/A', 'NA', 'nan', 'NaN', ''], np.nan
        )
        df['Top Post Impressions'] = pd.to_numeric(df['Top Post Impressions'], errors='coerce')
        
        # 4. Standardize post types
        df['Top Post Type'] = df['Top Post Type'].str.lower().str.replace(' ', '_')
        
        # 5. Ensure correct data types
        df['Week Starting'] = pd.to_datetime(df['Week Starting'])
        numeric_cols = ['Followers (Total)', 'New Followers', 'Impressions', 
                       'Engagements', 'Link Clicks']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        self.data['Social Media Clean'] = df
        self.log_step("   ✓ Social Media cleaned successfully")
        return self
    
    def clean_email_campaigns(self):
        """Clean Email Campaigns sheet"""
        self.log_step("📧 Cleaning Email Campaigns data...")
        df = self.data['Email Campaigns'].copy()
        
        # 1. Standardize campaign names
        df['Campaign Name'] = df['Campaign Name'].str.strip().str.title()
        
        # 2. Fix rate columns (Open Rate and CTR)
        def fix_rate_column(value):
            if pd.isna(value):
                return np.nan
            if isinstance(value, str):
                value = value.replace('%', '').strip()
                try:
                    num = float(value)
                    if num > 1:  # If > 1, likely percentage
                        return num / 100
                    return num
                except:
                    return np.nan
            elif isinstance(value, (int, float)):
                if value > 1:
                    return value / 100
                return value
            return np.nan
        
        df['Open Rate'] = df['Open Rate'].apply(fix_rate_column)
        df['CTR'] = df['CTR'].apply(fix_rate_column)
        
        # 3. Fill missing revenue with 0
        revenue_missing = df['Revenue Attributed'].isna().sum()
        df['Revenue Attributed'] = df['Revenue Attributed'].fillna(0)
        if revenue_missing > 0:
            self.log_step(f"   Filled {revenue_missing} missing revenue values with 0")
        
        # 4. Convert dates
        df['Send Date'] = pd.to_datetime(df['Send Date'])
        
        # 5. Ensure numeric columns
        numeric_cols = ['List Size', 'Emails Delivered', 'Opens', 'Clicks', 
                       'Unsubscribes', 'Conversions (Ticket Inquiries)', 'Revenue Attributed']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        self.data['Email Campaigns Clean'] = df
        self.log_step("   ✓ Email Campaigns cleaned successfully")
        return self
    
    def clean_sales_pipeline(self):
        """Clean Sales Pipeline sheet"""
        self.log_step("💰 Cleaning Sales Pipeline data...")
        df = self.data['Sales Pipeline'].copy()
        
        # 1. Standardize deal stages
        stage_mapping = {
            'Contacted': 'Contacted',
            'Lead': 'Lead',
            'Qualified': 'Qualified',
            'Negotiation': 'Negotiation',
            'Proposal Sent': 'Proposal Sent',
            'Closed Lost': 'Closed Lost',
            'Closed Won': 'Closed Won'
        }
        df['Deal Stage'] = df['Deal Stage'].replace(stage_mapping)
        
        # 2. Clean deal value (remove € symbol, commas)
        def clean_currency(value):
            if pd.isna(value):
                return 0.0
            if isinstance(value, str):
                # Remove currency symbols, commas, spaces
                value = value.replace('€', '').replace(',', '').strip()
                try:
                    return float(value)
                except:
                    return 0.0
            return float(value)
        
        df['Deal Value (EUR)'] = df['Deal Value (EUR)'].apply(clean_currency)
        
        # 3. Standardize lead source
        df['Lead Source'] = df['Lead Source'].str.title()
        
        # 4. Standardize ticket type
        df['Ticket Type'] = df['Ticket Type'].str.title()
        
        # 5. Fix dates
        date_columns = ['First Contact Date', 'Last Activity Date', 'Expected Close Date']
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # 6. Fill missing emails (simple pattern)
        def infer_email(row):
            if pd.isna(row['Contact Email']) or row['Contact Email'] == '':
                if pd.notna(row['Contact Name']) and pd.notna(row['Company Name']):
                    name_part = row['Contact Name'].split()[0].lower()
                    company_part = row['Company Name'].split()[0].lower().replace(' ', '')
                    return f"{name_part}@{company_part}.com"
            return row['Contact Email']
        
        missing_emails = df['Contact Email'].isna().sum()
        df['Contact Email'] = df.apply(infer_email, axis=1)
        if missing_emails > 0:
            self.log_step(f"   Inferred {missing_emails} missing emails")
        
        self.data['Sales Pipeline Clean'] = df
        self.log_step("   ✓ Sales Pipeline cleaned successfully")
        return self
    
    def clean_ad_spend(self):
        """Clean Ad Spend sheet"""
        self.log_step("📢 Cleaning Ad Spend data...")
        df = self.data['Ad Spend'].copy()
        
        # 1. Standardize month format
        df['Month'] = pd.to_datetime(df['Month'], format='%B %Y', errors='coerce')
        
        # 2. Clean currency columns
        currency_cols = ['Budget (EUR)', 'Spend (EUR)', 'CPM (EUR)', 'CPC (EUR)', 
                        'Cost per Conversion (EUR)']
        
        for col in currency_cols:
            # Remove € symbol and commas
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace('€', '', regex=False)
                df[col] = df[col].str.replace(',', '', regex=False)
            
            # Convert to numeric
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # 3. Calculate missing cost per conversion
        missing_cpc = df['Cost per Conversion (EUR)'].isna().sum()
        if missing_cpc > 0:
            mask = df['Cost per Conversion (EUR)'].isna() & (df['Conversions'] > 0)
            df.loc[mask, 'Cost per Conversion (EUR)'] = (
                df.loc[mask, 'Spend (EUR)'] / df.loc[mask, 'Conversions']
            )
            self.log_step(f"   Calculated {missing_cpc} missing cost per conversion values")
        
        # 4. Standardize campaign names
        df['Campaign Name'] = df['Campaign Name'].str.strip()
        
        self.data['Ad Spend Clean'] = df
        self.log_step("   ✓ Ad Spend cleaned successfully")
        return self
    
    def calculate_kpis(self):
        """Calculate key performance indicators"""
        self.log_step("📊 Calculating KPIs...")
        
        try:
            # Load cleaned data
            sales_df = self.data['Sales Pipeline Clean']
            ad_df = self.data['Ad Spend Clean']
            
            # 1. Basic metrics
            total_leads = len(sales_df)
            
            # Closed deals (won or lost)
            closed_deals = sales_df[sales_df['Deal Stage'].isin(['Closed Won', 'Closed Lost'])]
            conversion_rate = 0
            if len(closed_deals) > 0:
                conversion_rate = len(closed_deals[closed_deals['Deal Stage'] == 'Closed Won']) / len(closed_deals) * 100
            
            # 2. Revenue metrics
            won_deals = sales_df[sales_df['Deal Stage'] == 'Closed Won']
            total_revenue = won_deals['Deal Value (EUR)'].sum()
            total_pipeline = sales_df['Deal Value (EUR)'].sum()
            
            # 3. Marketing metrics
            total_ad_spend = ad_df['Spend (EUR)'].sum()
            total_conversions = ad_df['Conversions'].sum()
            overall_cpa = total_ad_spend / total_conversions if total_conversions > 0 else 0
            
            # 4. Progress toward targets
            delegate_target = 300
            sponsor_target = 25
            
            # Current progress (from closed won deals)
            current_delegates = len(sales_df[
                (sales_df['Ticket Type'].str.contains('Delegate', na=False)) & 
                (sales_df['Deal Stage'] == 'Closed Won')
            ])
            
            current_sponsors = len(sales_df[
                (sales_df['Ticket Type'].str.contains('Sponsor', na=False)) & 
                (sales_df['Deal Stage'] == 'Closed Won')
            ])
            
            # Calculate monthly growth
            monthly_leads = sales_df.groupby(
                sales_df['First Contact Date'].dt.to_period('M')
            ).size()
            monthly_growth = 0
            if len(monthly_leads) > 1:
                monthly_growth = monthly_leads.pct_change().mean() * 100
            
            kpis = {
                'total_leads': int(total_leads),
                'conversion_rate': round(conversion_rate, 1),
                'total_revenue': round(total_revenue, 2),
                'total_pipeline': round(total_pipeline, 2),
                'total_ad_spend': round(total_ad_spend, 2),
                'overall_cpa': round(overall_cpa, 2),
                'current_delegates': int(current_delegates),
                'current_sponsors': int(current_sponsors),
                'delegate_target': delegate_target,
                'sponsor_target': sponsor_target,
                'delegate_progress': round(current_delegates/delegate_target*100, 1) if delegate_target > 0 else 0,
                'sponsor_progress': round(current_sponsors/sponsor_target*100, 1) if sponsor_target > 0 else 0,
                'monthly_growth': round(monthly_growth, 1),
                'data_cleaned_on': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            self.data['KPIs'] = kpis
            self.log_step("   ✓ KPIs calculated successfully")
            return self
            
        except Exception as e:
            self.log_step(f"   ✗ Error calculating KPIs: {e}")
            # Return empty KPIs if error
            self.data['KPIs'] = {}
            return self
    
    def save_cleaned_data(self, output_dir='3_Cleaned_Data'):
        """Save all cleaned data to CSV files"""
        self.log_step(f"💾 Saving cleaned data to '{output_dir}'...")
        
        # Create directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Save each cleaned dataset
        cleaned_datasets = {
            'Website_Traffic_Clean': 'Website Traffic Clean',
            'Social_Media_Clean': 'Social Media Clean',
            'Email_Campaigns_Clean': 'Email Campaigns Clean',
            'Sales_Pipeline_Clean': 'Sales Pipeline Clean',
            'Ad_Spend_Clean': 'Ad Spend Clean'
        }
        
        for filename, dataset_name in cleaned_datasets.items():
            if dataset_name in self.data:
                filepath = os.path.join(output_dir, f"{filename}.csv")
                self.data[dataset_name].to_csv(filepath, index=False)
                self.log_step(f"   ✓ Saved {filename}.csv")
        
        # Save KPIs as JSON
        import json
        kpis_filepath = os.path.join(output_dir, "ml_insights.json")
        with open(kpis_filepath, 'w') as f:
            json.dump(self.data['KPIs'], f, indent=2, default=str)
        self.log_step(f"   ✓ Saved ml_insights.json")
        
        # Save cleaning log
        log_filepath = os.path.join(output_dir, "cleaning_log.txt")
        with open(log_filepath, 'w') as f:
            f.write("\n".join(self.cleaning_log))
        
        return self
    
    def run_full_pipeline(self):
        """Run complete data cleaning pipeline"""
        self.log_step("🚀 Starting complete data cleaning pipeline...")
        
        try:
            # Load all data first
            self.load_all_sheets()
            
            # Clean each dataset
            self.clean_website_traffic()
            self.clean_social_media()
            self.clean_email_campaigns()
            self.clean_sales_pipeline()
            self.clean_ad_spend()
            
            # Calculate KPIs
            self.calculate_kpis()
            
            # Save cleaned data
            self.save_cleaned_data()
            
            self.log_step("✅ Data cleaning pipeline completed successfully!")
            return True
            
        except Exception as e:
            self.log_step(f"❌ Pipeline failed with error: {e}")
            return False

# Main execution
if __name__ == "__main__":
    print("=" * 60)
    print("PROOF OF TALK 2026 - DATA CLEANING PIPELINE")
    print("=" * 60)
    
    # Initialize cleaner
    cleaner = POTDataCleaner('POT2026_Raw_Data_Case_Study.xlsx')
    
    # Run pipeline
    success = cleaner.run_full_pipeline()
    
    if success:
        print("\n" + "=" * 60)
        print("SUMMARY:")
        print("=" * 60)
        print(f"• Cleaned 5 data sources")
        print(f"• Generated KPIs for dashboard")
        print(f"• Saved files to '3_Cleaned_Data/' folder")
        print(f"• Cleaning log saved: 3_Cleaned_Data/cleaning_log.txt")
        print("=" * 60)
        
        # Display some sample KPIs
        if 'KPIs' in cleaner.data:
            kpis = cleaner.data['KPIs']
            print(f"\n📊 Sample KPIs Calculated:")
            print(f"  Total Leads: {kpis.get('total_leads', 'N/A')}")
            print(f"  Conversion Rate: {kpis.get('conversion_rate', 'N/A')}%")
            print(f"  Current Delegates: {kpis.get('current_delegates', 'N/A')}/{kpis.get('delegate_target', 'N/A')}")
            print(f"  Current Sponsors: {kpis.get('current_sponsors', 'N/A')}/{kpis.get('sponsor_target', 'N/A')}")
    else:
        print("\n❌ Pipeline failed. Check error messages above.")