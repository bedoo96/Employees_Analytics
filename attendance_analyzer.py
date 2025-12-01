import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re

class AttendanceAnalyzer:
    """
    Comprehensive attendance data analyzer with powerful processing capabilities
    """
    
    def __init__(self, file):
        """Initialize analyzer with uploaded file"""
        self.raw_df = pd.read_excel(file)
        self.df = self._process_dataframe()
        self.month_year = self._extract_month_year()
        
    def _process_dataframe(self):
        """Process and clean the attendance dataframe"""
        # The first row contains the actual column headers
        df = self.raw_df.copy()
        
        # Set first row as column names
        new_columns = df.iloc[0].tolist()
        df.columns = new_columns
        df = df.iloc[1:].reset_index(drop=True)
        
        # Clean column names
        df.columns = [str(col).strip() if col is not None else f'Col_{i}' 
                     for i, col in enumerate(df.columns)]
        
        # Convert numeric columns
        numeric_cols = ['Employee ID', 'Regular(H)', 'Late In(M)', 'Early Out(M)', 
                       'Absence(H)', 'Normal OT(H)', 'Weekend OT(H)', 'Holiday OT(H)',
                       'OT1(H)', 'OT2(H)', 'OT3(H)', 'Annual Leave(H)', 'Sick Leave(H)',
                       'Casual Leave(H)', 'Maternity Leave(H)', 'Compassionate Leave(H)',
                       'Business Trip(H)', 'Compensatory(H)', 'Compensatory Leave(H)']
        
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        return df
    
    def _extract_month_year(self):
        """Extract month and year from the data"""
        # Try to find date columns (numbered 1-31)
        date_cols = [col for col in self.df.columns if str(col).isdigit()]
        if date_cols:
            # Assume current year and month from file or current date
            return datetime.now().strftime('%B %Y')
        return datetime.now().strftime('%B %Y')
    
    def get_data_summary(self):
        """Get a comprehensive summary of the attendance data"""
        summary = {
            'total_employees': len(self.df),
            'departments': self.df['Department'].unique().tolist() if 'Department' in self.df.columns else [],
            'columns': self.df.columns.tolist(),
            'date_range': self.month_year,
            'metrics_available': {
                'working_hours': 'Regular(H)' in self.df.columns,
                'late_arrivals': 'Late In(M)' in self.df.columns,
                'early_departures': 'Early Out(M)' in self.df.columns,
                'overtime': 'Normal OT(H)' in self.df.columns,
                'leaves': any('Leave' in col for col in self.df.columns),
                'absences': 'Absence(H)' in self.df.columns
            },
            'sample_data': self.df.head(3).to_dict('records')
        }
        return summary
    
    def get_late_employees(self, min_times=1):
        """Get employees who came late more than specified times"""
        if 'Late In(M)' in self.df.columns:
            late_df = self.df[self.df['Late In(M)'] > 0].copy()
            
            # Count late days from daily columns
            date_cols = [col for col in self.df.columns if str(col).isdigit()]
            if date_cols:
                # This is simplified - actual implementation would check time stamps
                result = late_df[['Employee ID', 'First Name', 'Department', 'Late In(M)']].copy()
                result['Late_Count'] = (late_df['Late In(M)'] / 30).apply(np.ceil)  # Rough estimate
                return result[result['Late_Count'] >= min_times]
            else:
                return late_df[['Employee ID', 'First Name', 'Department', 'Late In(M)']]
        return pd.DataFrame()
    
    def get_overtime_analysis(self):
        """Analyze overtime hours"""
        ot_cols = ['Normal OT(H)', 'Weekend OT(H)', 'Holiday OT(H)']
        available_ot = [col for col in ot_cols if col in self.df.columns]
        
        if available_ot:
            result = self.df[['Employee ID', 'First Name', 'Department'] + available_ot].copy()
            result['Total OT'] = result[available_ot].sum(axis=1)
            return result.sort_values('Total OT', ascending=False)
        return pd.DataFrame()
    
    def get_working_hours_summary(self):
        """Get working hours summary for all employees"""
        if 'Regular(H)' in self.df.columns:
            result = self.df[['Employee ID', 'First Name', 'Department', 'Regular(H)']].copy()
            result = result.sort_values('Regular(H)', ascending=False)
            return result
        return pd.DataFrame()
    
    def get_leave_analysis(self):
        """Analyze leave patterns"""
        leave_cols = [col for col in self.df.columns if 'Leave' in col]
        
        if leave_cols:
            result = self.df[['Employee ID', 'First Name', 'Department'] + leave_cols].copy()
            result['Total Leave(H)'] = result[leave_cols].sum(axis=1)
            return result[result['Total Leave(H)'] > 0].sort_values('Total Leave(H)', ascending=False)
        return pd.DataFrame()
    
    def get_absence_analysis(self):
        """Analyze absence patterns"""
        if 'Absence(H)' in self.df.columns:
            result = self.df[self.df['Absence(H)'] > 0][['Employee ID', 'First Name', 
                                                          'Department', 'Absence(H)']].copy()
            return result.sort_values('Absence(H)', ascending=False)
        return pd.DataFrame()
    
    def get_department_summary(self):
        """Get department-wise summary"""
        if 'Department' not in self.df.columns:
            return pd.DataFrame()
        
        summary_cols = {
            'Regular(H)': 'sum',
            'Late In(M)': 'sum',
            'Early Out(M)': 'sum',
            'Normal OT(H)': 'sum',
            'Absence(H)': 'sum'
        }
        
        available_cols = {k: v for k, v in summary_cols.items() if k in self.df.columns}
        
        if available_cols:
            dept_summary = self.df.groupby('Department').agg({
                'Employee ID': 'count',
                **available_cols
            }).reset_index()
            dept_summary.rename(columns={'Employee ID': 'Employee_Count'}, inplace=True)
            return dept_summary
        return pd.DataFrame()
    
    def get_top_performers(self, n=10, metric='Regular(H)'):
        """Get top N employees by specified metric"""
        if metric in self.df.columns:
            return self.df.nlargest(n, metric)[['Employee ID', 'First Name', 
                                                 'Department', metric]]
        return pd.DataFrame()
    
    def get_weekly_breakdown(self):
        """Get weekly breakdown of attendance (if daily data available)"""
        # This would require daily timestamp columns
        # Simplified version based on total hours
        if 'Regular(H)' in self.df.columns:
            result = self.df[['Employee ID', 'First Name', 'Department', 'Regular(H)']].copy()
            # Estimate weekly hours (assuming ~4.33 weeks per month)
            result['Avg Weekly Hours'] = (result['Regular(H)'] / 4.33).round(2)
            return result
        return pd.DataFrame()
    
    def get_punctuality_score(self):
        """Calculate punctuality score for employees"""
        result = self.df[['Employee ID', 'First Name', 'Department']].copy()
        
        # Calculate score based on late arrivals and early departures
        if 'Late In(M)' in self.df.columns and 'Early Out(M)' in self.df.columns:
            total_violations = self.df['Late In(M)'] + self.df['Early Out(M)']
            # Score out of 100 (higher is better)
            result['Punctuality Score'] = (100 - (total_violations / 10)).clip(0, 100).round(2)
        elif 'Late In(M)' in self.df.columns:
            result['Punctuality Score'] = (100 - (self.df['Late In(M)'] / 10)).clip(0, 100).round(2)
        else:
            result['Punctuality Score'] = 100
        
        return result.sort_values('Punctuality Score', ascending=False)
    
    def generate_monthly_report(self, employee_id=None):
        """Generate comprehensive monthly report"""
        if employee_id:
            employee_data = self.df[self.df['Employee ID'] == employee_id]
        else:
            employee_data = self.df
        
        report = {
            'total_employees': len(employee_data),
            'total_working_hours': employee_data['Regular(H)'].sum() if 'Regular(H)' in self.df.columns else 0,
            'total_overtime': employee_data['Normal OT(H)'].sum() if 'Normal OT(H)' in self.df.columns else 0,
            'total_late_minutes': employee_data['Late In(M)'].sum() if 'Late In(M)' in self.df.columns else 0,
            'total_absences': employee_data['Absence(H)'].sum() if 'Absence(H)' in self.df.columns else 0,
            'avg_working_hours': employee_data['Regular(H)'].mean() if 'Regular(H)' in self.df.columns else 0,
        }
        
        return report
    
    def search_employees(self, query):
        """Search employees by name or ID"""
        query = str(query).lower()
        mask = (
            self.df['First Name'].astype(str).str.lower().str.contains(query, na=False) |
            self.df['Employee ID'].astype(str).str.contains(query, na=False)
        )
        if 'Department' in self.df.columns:
            mask |= self.df['Department'].astype(str).str.lower().str.contains(query, na=False)
        
        return self.df[mask]