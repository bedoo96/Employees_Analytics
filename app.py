import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import json
from attendance_analyzer import AttendanceAnalyzer
from llm_handler import LLMHandler

# Page configuration
st.set_page_config(
    page_title="Employee Attendance Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .suggestion-box {
        background-color: #e8f4f8;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .stButton>button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'analyzer' not in st.session_state:
    st.session_state.analyzer = None
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []

# Sidebar - Company Logo and Info
with st.sidebar:
    # Company Logo Placeholder
    st.image("https://via.placeholder.com/200x80/1f77b4/ffffff?text=Company+Logo", 
             use_container_width=True)
    st.markdown("---")
    
    st.markdown("### 🤖 AI Model Settings")
    
    # Initialize LLM Handler
    llm_handler = LLMHandler()
    
    # Model selection
    model_type = st.radio(
        "Select AI Provider:",
        ["OpenAI", "Ollama (Local)"],
        help="Choose between cloud-based OpenAI or local Ollama models"
    )
    
    if model_type == "OpenAI":
        api_key = st.text_input(
            "OpenAI API Key:",
            type="password",
            help="Enter your OpenAI API key"
        )
        if api_key:
            llm_handler.set_openai_key(api_key)
        selected_model = st.selectbox(
            "Model:",
            ["gpt-4", "gpt-3.5-turbo"],
            help="Select OpenAI model"
        )
    else:
        # Check available Ollama models
        available_models = llm_handler.get_ollama_models()
        if available_models:
            selected_model = st.selectbox(
                "Available Ollama Models:",
                available_models,
                help="Models available on your local machine"
            )
        else:
            st.warning("⚠️ No Ollama models found. Please install Ollama and pull a model.")
            st.code("ollama pull llama2")
            selected_model = None
    
    if selected_model:
        llm_handler.initialize_model(model_type, selected_model)
        st.success(f"✅ {selected_model} ready")
    
    st.markdown("---")
    
    # File upload section
    st.markdown("### 📁 Upload Attendance File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload monthly employee check-in/check-out records"
    )
    
    if uploaded_file:
        try:
            analyzer = AttendanceAnalyzer(uploaded_file)
            st.session_state.df = analyzer.df
            st.session_state.analyzer = analyzer
            st.success(f"✅ File loaded: {len(analyzer.df)} employees")
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")

# Main content
st.markdown('<p class="main-header">📊 Employee Attendance Analytics System</p>', 
            unsafe_allow_html=True)
st.markdown('<p class="sub-header">AI-Powered Workforce Management & Reporting</p>', 
            unsafe_allow_html=True)

# Description section
with st.expander("ℹ️ About This Application", expanded=False):
    st.markdown("""
    This intelligent system analyzes employee attendance data and generates comprehensive reports using AI.
    
    **Key Features:**
    - 📈 Automated analysis of check-in/check-out records
    - 🤖 AI-powered natural language query processing
    - 📊 Interactive visualizations and dashboards
    - 📝 Generate custom reports based on your questions
    - ⚡ Real-time insights on employee attendance patterns
    
    **What you can analyze:**
    - Weekly/Monthly working hours per employee
    - Late arrivals and early departures tracking
    - Overtime hours calculation
    - Leave analysis (annual, sick, casual, etc.)
    - Absence patterns and trends
    - Department-wise performance metrics
    """)

# Main interface
if st.session_state.df is not None:
    analyzer = st.session_state.analyzer
    
    # Quick Stats Dashboard
    st.markdown("### 📊 Quick Overview")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Employees", len(analyzer.df))
    with col2:
        avg_hours = analyzer.df['Regular(H)'].mean() if 'Regular(H)' in analyzer.df.columns else 0
        st.metric("Avg Working Hours", f"{avg_hours:.1f}h")
    with col3:
        total_late = analyzer.df['Late In(M)'].sum() if 'Late In(M)' in analyzer.df.columns else 0
        st.metric("Total Late Arrivals", f"{int(total_late)} min")
    with col4:
        total_ot = analyzer.df['Normal OT(H)'].sum() if 'Normal OT(H)' in analyzer.df.columns else 0
        st.metric("Total Overtime", f"{total_ot:.1f}h")
    
    st.markdown("---")
    
    # Data Preview
    with st.expander("📋 View Attendance Data", expanded=False):
        st.dataframe(analyzer.df, use_container_width=True, height=400)
    
    st.markdown("---")
    
    # AI Query Interface
    st.markdown("### 💬 Ask Questions About Your Data")
    
    # Suggested questions
    st.markdown('<div class="suggestion-box">', unsafe_allow_html=True)
    st.markdown("**💡 Suggested Questions:**")
    suggestions = [
        "Show me employees who came late more than 5 times",
        "Calculate total working hours for each employee this month",
        "Which employees have the most overtime hours?",
        "Generate a weekly report for all employees",
        "Show me absence patterns by department",
        "Who took the most annual leave?",
        "Calculate average early departure time",
        "Show me weekend overtime distribution"
    ]
    
    cols = st.columns(2)
    for idx, suggestion in enumerate(suggestions):
        col = cols[idx % 2]
        if col.button(suggestion, key=f"sug_{idx}"):
            st.session_state.current_query = suggestion
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Custom query input
    user_query = st.text_area(
        "Or type your custom question:",
        value=st.session_state.get('current_query', ''),
        height=100,
        placeholder="E.g., 'Show me the top 5 employees with highest working hours' or 'Generate a monthly summary report'"
    )
    
    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        generate_button = st.button("🚀 Generate Report", type="primary")
    with col2:
        clear_button = st.button("🗑️ Clear History")
    
    if clear_button:
        st.session_state.chat_history = []
        st.rerun()
    
    # Process query
    if generate_button and user_query and llm_handler.llm:
        with st.spinner("🤖 AI is analyzing your data..."):
            try:
                # Get data summary for context
                data_summary = analyzer.get_data_summary()
                
                # Generate response using LLM
                response = llm_handler.generate_report(user_query, data_summary, analyzer.df)
                
                # Add to chat history
                st.session_state.chat_history.append({
                    'query': user_query,
                    'response': response,
                    'timestamp': datetime.now()
                })
                
            except Exception as e:
                st.error(f"❌ Error generating report: {str(e)}")
    
    # Display chat history
    if st.session_state.chat_history:
        st.markdown("---")
        st.markdown("### 📜 Analysis History")
        
        for idx, chat in enumerate(reversed(st.session_state.chat_history)):
            with st.container():
                st.markdown(f"**🙋 Question ({chat['timestamp'].strftime('%H:%M:%S')}):**")
                st.info(chat['query'])
                st.markdown("**🤖 Analysis:**")
                st.markdown(chat['response'])
                
                # Try to execute any analysis code
                if '```python' in chat['response']:
                    code_blocks = chat['response'].split('```python')
                    for block in code_blocks[1:]:
                        code = block.split('```')[0].strip()
                        try:
                            # Execute visualization code safely
                            exec(code, {'df': analyzer.df, 'st': st, 'pd': pd, 
                                       'px': px, 'go': go, 'analyzer': analyzer})
                        except Exception as e:
                            st.warning(f"Could not execute visualization: {str(e)}")
                
                st.markdown("---")
    
    elif not llm_handler.llm:
        st.warning("⚠️ Please configure an AI model in the sidebar to start analyzing.")

else:
    # Welcome screen
    st.markdown("### 👋 Welcome!")
    st.info("""
    **Getting Started:**
    1. 📁 Upload your employee attendance Excel file using the sidebar
    2. 🤖 Configure your preferred AI model (OpenAI or local Ollama)
    3. 💬 Ask questions about your data in natural language
    4. 📊 Get instant insights and visualizations
    
    **Expected File Format:**
    Your Excel file should contain employee check-in/check-out records with columns like:
    - Employee ID, Name, Department
    - Daily attendance (dates as columns)
    - Summary metrics (Regular hours, Late In, Early Out, Overtime, Leaves, etc.)
    """)
    
    # Sample visualization
    st.markdown("### 📈 Sample Analytics Dashboard")
    
    # Create sample charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Sample attendance trend
        sample_dates = pd.date_range('2024-11-01', '2024-11-30', freq='D')
        sample_data = pd.DataFrame({
            'Date': sample_dates,
            'Attendance': [85 + (i % 10) for i in range(len(sample_dates))]
        })
        fig = px.line(sample_data, x='Date', y='Attendance', 
                     title='Daily Attendance Rate (%)')
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Sample department distribution
        sample_dept = pd.DataFrame({
            'Department': ['Engineering', 'Sales', 'HR', 'Operations'],
            'Employees': [45, 32, 12, 28]
        })
        fig = px.pie(sample_dept, values='Employees', names='Department',
                    title='Employee Distribution by Department')
        st.plotly_chart(fig, use_container_width=True)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 2rem;'>
    <p>🚀 Powered by AI | Built with Streamlit, LangChain & OpenAI/Ollama</p>
    <p style='font-size: 0.9rem;'>© 2024 Employee Attendance Analytics System</p>
</div>
""", unsafe_allow_html=True)