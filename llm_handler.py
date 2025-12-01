from langchain_openai import ChatOpenAI
from langchain_community.llms import Ollama
# from langchain.prompts import PromptTemplate
# from langchain.chains import LLMChain
from langchain_core.prompts import PromptTemplate
# Remove the LLMChain import - we'll use invoke directly
import requests
import json
import pandas as pd
import traceback

class LLMHandler:
    """
    Handler for integrating OpenAI and Ollama models using LangChain
    """
    
    def __init__(self):
        self.llm = None
        self.model_type = None
        self.model_name = None
        self.openai_api_key = None
        
    def set_openai_key(self, api_key):
        """Set OpenAI API key"""
        self.openai_api_key = api_key
        
    def get_ollama_models(self):
        """Get list of available Ollama models on local machine"""
        try:
            response = requests.get('http://localhost:11434/api/tags', timeout=5)
            if response.status_code == 200:
                models_data = response.json()
                models = [model['name'] for model in models_data.get('models', [])]
                return models if models else None
            return None
        except requests.exceptions.RequestException:
            return None
    
    def initialize_model(self, model_type, model_name):
        """Initialize the selected LLM model"""
        self.model_type = model_type
        self.model_name = model_name
        
        try:
            if model_type == "OpenAI":
                if not self.openai_api_key:
                    raise ValueError("OpenAI API key is required")
                self.llm = ChatOpenAI(
                    model=model_name,
                    openai_api_key=self.openai_api_key,
                    temperature=0.3
                )
            elif model_type == "Ollama (Local)":
                self.llm = Ollama(
                    model=model_name,
                    base_url="http://localhost:11434"
                )
            return True
        except Exception as e:
            print(f"Error initializing model: {str(e)}")
            return False
    
    def create_analysis_prompt(self, query, data_summary, df_info):
        """Create a comprehensive prompt for data analysis"""
        
        template = """You are an expert HR data analyst. You have access to employee attendance data and need to answer questions about it.

**Data Summary:**
- Total Employees: {total_employees}
- Departments: {departments}
- Date Range: {date_range}
- Available Metrics: {metrics}

**Available Columns:**
{columns}

**Sample Data:**
{sample_data}

**User Question:**
{query}

**Instructions:**
1. Analyze the question carefully
2. Determine what data analysis or calculation is needed
3. Provide clear, actionable insights
4. Include specific numbers and statistics
5. Format your response in a clear, professional manner

**Important:**
- Be specific with numbers and percentages
- If the question requires sorting or filtering, explain the criteria
- Suggest actionable recommendations when relevant
- If you need to show data, format it as a clear table or list

Please provide a comprehensive answer:
"""
    # 4. If visualization would help, suggest appropriate chart types    
        prompt = PromptTemplate(
            input_variables=["total_employees", "departments", "date_range", 
                           "metrics", "columns", "sample_data", "query"],
            template=template
        )
        
        return prompt
    
    def generate_report(self, query, data_summary, df):
        """Generate a comprehensive report based on the query"""
        
        if not self.llm:
            return "❌ No LLM model initialized. Please configure a model in the sidebar."
        
        try:
            # Prepare data summary
            metrics_str = "\n".join([f"  - {k}: {'✓' if v else '✗'}" 
                                    for k, v in data_summary['metrics_available'].items()])
            
            columns_str = ", ".join(data_summary['columns'][:20]) + "..."
            
            sample_data_str = pd.DataFrame(data_summary['sample_data']).to_string()
            
            departments_str = ", ".join(data_summary['departments']) if data_summary['departments'] else "Not specified"
            
            # Create prompt
            prompt = self.create_analysis_prompt(
                query=query,
                data_summary=data_summary,
                df_info=df.head().to_string()
            )
            
            # Format the prompt
            formatted_prompt = prompt.format(
                total_employees=data_summary['total_employees'],
                departments=departments_str,
                date_range=data_summary['date_range'],
                metrics=metrics_str,
                columns=columns_str,
                sample_data=sample_data_str,
                query=query
            )
            
            # Get response from LLM
            if self.model_type == "OpenAI":
                response = self.llm.invoke(formatted_prompt)
                answer = response.content if hasattr(response, 'content') else str(response)
            else:
                response = self.llm.invoke(formatted_prompt)
                answer = response if isinstance(response, str) else str(response)
            
            # Enhance the response with actual data analysis
            enhanced_answer = self._enhance_with_data_analysis(query, answer, df)
            
            return enhanced_answer
            
        except Exception as e:
            error_trace = traceback.format_exc()
            return f"❌ Error generating report: {str(e)}\n\nDetails:\n{error_trace}"
    
    def _enhance_with_data_analysis(self, query, llm_answer, df):
        """Enhance LLM answer with actual data analysis and visualizations"""
        
        query_lower = query.lower()
        enhanced = llm_answer + "\n\n"
        
        # Add relevant data analysis based on query keywords
        
        if 'late' in query_lower or 'tardy' in query_lower:
            if 'Late In(M)' in df.columns:
                late_employees = df[df['Late In(M)'] > 0].nlargest(10, 'Late In(M)')
                enhanced += "\n**📊 Top 10 Late Arrivals:**\n\n"
                enhanced += late_employees[['Employee ID', 'First Name', 'Department', 'Late In(M)']].to_markdown(index=False)
                
                enhanced += "\n\n**📈 Visualization Code:**\n```python\nimport plotly.express as px\nlate_data = df[df['Late In(M)'] > 0].nlargest(10, 'Late In(M)')\nfig = px.bar(late_data, x='First Name', y='Late In(M)', \n             title='Top 10 Employees with Most Late Minutes',\n             labels={'Late In(M)': 'Late Minutes', 'First Name': 'Employee'},\n             color='Late In(M)', color_continuous_scale='Reds')\nst.plotly_chart(fig, use_container_width=True)\n```"
        
        if 'overtime' in query_lower or 'ot' in query_lower:
            if 'Normal OT(H)' in df.columns:
                ot_employees = df.nlargest(10, 'Normal OT(H)')
                enhanced += "\n**📊 Top 10 Overtime Workers:**\n\n"
                enhanced += ot_employees[['Employee ID', 'First Name', 'Department', 'Normal OT(H)']].to_markdown(index=False)
                
                enhanced += "\n\n**📈 Visualization Code:**\n```python\nimport plotly.express as px\not_data = df.nlargest(10, 'Normal OT(H)')\nfig = px.bar(ot_data, x='First Name', y='Normal OT(H)',\n             title='Top 10 Employees by Overtime Hours',\n             labels={'Normal OT(H)': 'Overtime (Hours)', 'First Name': 'Employee'},\n             color='Normal OT(H)', color_continuous_scale='Blues')\nst.plotly_chart(fig, use_container_width=True)\n```"
        
        if 'working hours' in query_lower or 'total hours' in query_lower:
            if 'Regular(H)' in df.columns:
                top_workers = df.nlargest(10, 'Regular(H)')
                enhanced += "\n**📊 Top 10 by Working Hours:**\n\n"
                enhanced += top_workers[['Employee ID', 'First Name', 'Department', 'Regular(H)']].to_markdown(index=False)
                
                avg_hours = df['Regular(H)'].mean()
                enhanced += f"\n\n**📈 Statistics:**\n- Average Working Hours: {avg_hours:.2f}h\n- Total Working Hours: {df['Regular(H)'].sum():.2f}h\n"
                
                enhanced += "\n**📊 Visualization Code:**\n```python\nimport plotly.express as px\ntop_data = df.nlargest(15, 'Regular(H)')\nfig = px.bar(top_data, x='First Name', y='Regular(H)',\n             title='Top 15 Employees by Working Hours',\n             labels={'Regular(H)': 'Working Hours', 'First Name': 'Employee'},\n             color='Regular(H)', color_continuous_scale='Greens')\nst.plotly_chart(fig, use_container_width=True)\n```"
        
        if 'leave' in query_lower or 'absence' in query_lower:
            leave_cols = [col for col in df.columns if 'Leave' in col]
            if leave_cols:
                df_temp = df.copy()
                df_temp['Total Leave'] = df_temp[leave_cols].sum(axis=1)
                top_leaves = df_temp[df_temp['Total Leave'] > 0].nlargest(10, 'Total Leave')
                enhanced += "\n**📊 Top 10 by Leave Hours:**\n\n"
                enhanced += top_leaves[['Employee ID', 'First Name', 'Department', 'Total Leave']].to_markdown(index=False)
        
        if 'department' in query_lower or 'dept' in query_lower:
            if 'Department' in df.columns and 'Regular(H)' in df.columns:
                dept_summary = df.groupby('Department').agg({
                    'Employee ID': 'count',
                    'Regular(H)': 'sum'
                }).reset_index()
                dept_summary.columns = ['Department', 'Employee Count', 'Total Hours']
                enhanced += "\n**📊 Department Summary:**\n\n"
                enhanced += dept_summary.to_markdown(index=False)
                
                enhanced += "\n\n**📈 Visualization Code:**\n```python\nimport plotly.express as px\ndept_data = df.groupby('Department').agg({'Employee ID': 'count', 'Regular(H)': 'sum'}).reset_index()\ndept_data.columns = ['Department', 'Employees', 'Total_Hours']\nfig = px.pie(dept_data, values='Employees', names='Department',\n             title='Employee Distribution by Department')\nst.plotly_chart(fig, use_container_width=True)\n```"
        
        if 'weekly' in query_lower or 'week' in query_lower:
            if 'Regular(H)' in df.columns:
                df_temp = df.copy()
                df_temp['Weekly Hours'] = (df_temp['Regular(H)'] / 4.33).round(2)
                enhanced += "\n**📊 Weekly Hours Breakdown (Estimated):**\n\n"
                enhanced += df_temp[['Employee ID', 'First Name', 'Regular(H)', 'Weekly Hours']].head(10).to_markdown(index=False)
                enhanced += f"\n\n*Note: Weekly hours estimated by dividing monthly hours by 4.33 weeks*"
        
        return enhanced