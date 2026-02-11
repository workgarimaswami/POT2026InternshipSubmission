# Proof of Talk 2026 - Technical Documentation

## 🏗️ Architecture Overview
POT2026_Data_Pipeline/
├── Data Extraction (Excel Files)
├── Data Cleaning (Python Pipeline)
├── ML Analysis (Scikit-learn, XGBoost)
├── Dashboard (Streamlit + Plotly)
└── Deployment (Streamlit Cloud)


## 📊 Data Pipeline

### 1. Data Cleaning (`1_data_cleaning_pipeline.py`)
**Input:** `POT2026_Raw_Data_Case_Study.xlsx` (5 sheets)
**Output:** Cleaned CSV files in `3_Cleaned_Data/`

**Cleaning Steps:**
- Remove duplicates
- Fix data types
- Handle missing values
- Standardize formats
- Calculate derived metrics

### 2. ML Analysis (`2_ml_analysis.py`)
**Models Used:**
- **Random Forest Regressor** for forecasting
- **K-Means Clustering** for hidden insights
- **ROI Calculation** for channel performance
- **Conversion Analysis** by lead source

**Output:** `ml_insights.json` with all analysis results

### 3. Dashboard (`3_streamlit_dashboard.py`)
**Features:**
- CEO 30-second view
- Interactive visualizations
- Real-time KPIs
- Downloadable reports
- Mobile responsive

## 🚀 Deployment Instructions

### Local Development
```bash
# 1. Create virtual environment
python -m venv venv

# 2. Activate virtual environment
# On Windows:
venv\Scripts\activate
# On Mac/Linux:
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run data pipeline
python 1_data_cleaning_pipeline.py

# 5. Run ML analysis
python 2_ml_analysis.py

# 6. Run dashboard
streamlit run 3_streamlit_dashboard.py