import pandas as pd
import datetime
import os
from typing import List, Dict, Tuple, Optional
import logging
from collections import defaultdict
import argparse

# Set up logging for better debugging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('attendance_processor.log'),
        logging.StreamHandler()
    ]
)

class AttendanceReportGenerator:
    """
    Enhanced Attendance Report Generator with improved Excel formatting
    """
    
    def __init__(self):
        self.df = None
        self.available_months = []
        self.employee_data = {}
        
    def read_attendance_file(self, input_file: str) -> bool:
        """Read and parse the attendance file with improved error handling"""
        print("--- Starting Attendance File Processing ---")
        logging.info(f"Attempting to read file: {input_file}")
        
        try:
            # Define column names as per the original structure
            col_names = ['No', 'TMNo', 'EnNo', 'Name', 'GMNo', 'Mode', 'IN/OUT', 'Antipass', 'DaiGong', 'DateTime', 'TR']
            
            # Read with multiple encoding attempts
            encodings_to_try = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252']
            
            for encoding in encodings_to_try:
                try:
                    self.df = pd.read_csv(
                        input_file, 
                        sep='\t', 
                        header=None, 
                        names=col_names, 
                        skiprows=5, 
                        encoding=encoding,
                        on_bad_lines='skip'
                    )
                    print(f"✅ Successfully read with {encoding} encoding")
                    logging.info(f"File read successfully with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
                    
            if self.df is None:
                raise Exception("Could not read file with any encoding")
                
            print(f"📊 Raw records read: {len(self.df)}")
            logging.info(f"Raw records count: {len(self.df)}")
            
            return True
            
        except FileNotFoundError:
            print(f"❌ File not found: '{input_file}'")
            logging.error(f"File not found: {input_file}")
            return False
        except Exception as e:
            print(f"❌ Error reading file: {str(e)}")
            logging.error(f"Error reading file: {str(e)}")
            return False
    
    def preprocess_data(self) -> bool:
        """Clean and preprocess the attendance data with debugging"""
        if self.df is None:
            return False
            
        print("\n🔧 Preprocessing data...")
        logging.info("Starting data preprocessing")
        
        # Store original count for comparison
        original_count = len(self.df)
        
        # Clean string columns
        string_columns = ['No', 'TMNo', 'EnNo', 'Name', 'GMNo', 'Mode', 'IN/OUT', 'Antipass', 'DaiGong', 'TR']
        for col in string_columns:
            if col in self.df.columns:
                self.df[col] = self.df[col].astype(str).str.strip()
        
        # Rename columns for clarity
        self.df.rename(columns={'EnNo': 'EmpID', 'Name': 'EmployeeName'}, inplace=True)
        
        # Remove rows with missing DateTime
        self.df.dropna(subset=['DateTime'], inplace=True)
        print(f"📉 Removed {original_count - len(self.df)} rows with missing DateTime")
        
        # Parse DateTime with debugging
        print("🕐 Parsing DateTime...")
        datetime_errors = []
        
        def parse_datetime_safe(dt_str):
            """Safely parse datetime with error tracking"""
            try:
                if pd.isna(dt_str) or str(dt_str).strip() == '':
                    return pd.NaT
                
                # Try multiple datetime formats
                formats_to_try = [
                    '%Y-%m-%d %H:%M:%S',
                    '%d/%m/%Y %H:%M:%S',
                    '%m/%d/%Y %H:%M:%S',
                    '%Y/%m/%d %H:%M:%S',
                    '%d-%m-%Y %H:%M:%S',
                    '%Y-%m-%d %H:%M',
                    '%d/%m/%Y %H:%M',
                ]
                
                dt_str_clean = str(dt_str).strip()
                
                for fmt in formats_to_try:
                    try:
                        return pd.to_datetime(dt_str_clean, format=fmt)
                    except:
                        continue
                
                # If all formats fail, try pandas auto-parsing
                return pd.to_datetime(dt_str_clean, errors='coerce')
                
            except Exception as e:
                datetime_errors.append(f"Error parsing '{dt_str}': {str(e)}")
                return pd.NaT
        
        self.df['DateTime'] = self.df['DateTime'].apply(parse_datetime_safe)
        
        # Remove rows with failed datetime parsing
        before_datetime_filter = len(self.df)
        self.df.dropna(subset=['DateTime'], inplace=True)
        datetime_failures = before_datetime_filter - len(self.df)
        
        if datetime_failures > 0:
            print(f"⚠️  Failed to parse {datetime_failures} datetime entries")
            logging.warning(f"DateTime parsing failures: {datetime_failures}")
        
        # Format EmpID consistently
        self.df['EmpID'] = self.df['EmpID'].astype(str).str.zfill(8)
        
        # Create YearMonth column for filtering
        self.df['YearMonth'] = self.df['DateTime'].dt.strftime('%Y-%m')
        
        # Get unique months and sort them
        self.available_months = sorted(self.df['YearMonth'].dropna().unique())
        
        print(f"✅ Data preprocessing complete!")
        print(f"📊 Final record count: {len(self.df)}")
        print(f"📅 Available months: {', '.join(self.available_months)}")
        print(f"👥 Unique employees: {self.df['EmpID'].nunique()}")
        
        logging.info(f"Preprocessing complete. Final records: {len(self.df)}")
        return True
    
    def get_month_selection_menu(self) -> List[str]:
        """Interactive month selection using hash-based lookup"""
        if not self.available_months:
            print("❌ No months available for selection")
            return []
        
        # Create a hash map for quick month lookup
        month_lookup = {str(i+1): month for i, month in enumerate(self.available_months)}
        month_names = {}
        
        print("\n📅 Available Months for Report Generation:")
        print("=" * 50)
        
        for i, month in enumerate(self.available_months, 1):
            try:
                # Convert to readable format
                year, month_num = month.split('-')
                month_name = datetime.datetime(int(year), int(month_num), 1).strftime('%B %Y')
                month_names[month] = month_name
                
                # Count records for this month
                month_records = len(self.df[self.df['YearMonth'] == month])
                employees_count = self.df[self.df['YearMonth'] == month]['EmpID'].nunique()
                
                print(f"{i:2d}. {month_name:<15} ({month_records:4d} records, {employees_count:2d} employees)")
            except:
                print(f"{i:2d}. {month}")
        
        print("\n" + "=" * 50)
        print("Selection Options:")
        print("• Enter numbers (e.g., 1,3,5 for months 1, 3, and 5)")
        print("• Enter 'all' to select all months")
        print("• Enter ranges (e.g., 1-3 for months 1 to 3)")
        print("• Enter 'quit' to exit")
        
        while True:
            try:
                selection = input("\n🎯 Select months: ").strip().lower()
                
                if selection == 'quit':
                    return []
                
                if selection == 'all':
                    return self.available_months.copy()
                
                # Parse selection using set for efficient operations
                selected_indices = set()
                
                # Split by comma and process each part
                parts = [part.strip() for part in selection.split(',')]
                
                for part in parts:
                    if '-' in part and part.count('-') == 1:
                        # Handle range (e.g., "1-3")
                        try:
                            start, end = map(int, part.split('-'))
                            if 1 <= start <= len(self.available_months) and 1 <= end <= len(self.available_months):
                                selected_indices.update(range(start, end + 1))
                            else:
                                print(f"⚠️  Invalid range: {part}")
                        except ValueError:
                            print(f"⚠️  Invalid range format: {part}")
                    else:
                        # Handle single number
                        try:
                            num = int(part)
                            if 1 <= num <= len(self.available_months):
                                selected_indices.add(num)
                            else:
                                print(f"⚠️  Invalid month number: {num}")
                        except ValueError:
                            print(f"⚠️  Invalid input: {part}")
                
                if selected_indices:
                    # Convert indices to actual months using our hash map
                    selected_months = [month_lookup[str(i)] for i in sorted(selected_indices)]
                    
                    print(f"\n✅ Selected months:")
                    for month in selected_months:
                        print(f"   • {month_names.get(month, month)}")
                    
                    confirm = input("\n❓ Confirm selection? (y/n): ").strip().lower()
                    if confirm in ['y', 'yes']:
                        return selected_months
                    
                else:
                    print("❌ No valid months selected. Please try again.")
                    
            except KeyboardInterrupt:
                print("\n\n👋 Goodbye!")
                return []
            except Exception as e:
                print(f"❌ Error in selection: {str(e)}")
    
    def generate_monthly_report(self, month: str) -> Optional[pd.DataFrame]:
        """Generate report for a specific month with simple row format per employee"""
        print(f"\n📊 Processing: {month}")
        logging.info(f"Generating report for month: {month}")
        
        # Filter data for the specific month
        month_df = self.df[self.df['YearMonth'] == month].copy()
        
        if month_df.empty:
            print(f"⚠️  No data found for {month}")
            return None
        
        # Get date range for the month
        year, month_num = month.split('-')
        start_date = datetime.datetime(int(year), int(month_num), 1)
        
        # Calculate last day of month
        if int(month_num) == 12:
            end_date = datetime.datetime(int(year) + 1, 1, 1) - datetime.timedelta(days=1)
        else:
            end_date = datetime.datetime(int(year), int(month_num) + 1, 1) - datetime.timedelta(days=1)
        
        # Generate all dates in the month
        all_days = pd.date_range(start=start_date, end=end_date, freq='D')
        
        # Get unique employees for this month
        employees = sorted(month_df[['EmpID', 'EmployeeName']].drop_duplicates().values.tolist())
        
        print(f"   📈 Processing {len(employees)} employees for {len(all_days)} days")
        
        # Initialize report data list
        report_rows = []
        
        # Process each employee
        for emp_id, emp_name in employees:
            emp_records = month_df[month_df['EmpID'] == emp_id].copy()
            
            # Separate IN and OUT records with better pattern matching
            in_patterns = ['time in', 'in', 'entry', 'check in']
            out_patterns = ['time out', 'out', 'exit', 'check out']
            
            df_in = emp_records[emp_records['TR'].str.lower().str.contains('|'.join(in_patterns), na=False)]
            df_out = emp_records[emp_records['TR'].str.lower().str.contains('|'.join(out_patterns), na=False)]
            
            # Group by date and get earliest IN and latest OUT
            in_times = df_in.groupby(df_in['DateTime'].dt.date)['DateTime'].min()
            out_times = df_out.groupby(df_out['DateTime'].dt.date)['DateTime'].max()
            
            # Create employee header row
            employee_header = {
                'Employee_Info': f"{emp_id} - {emp_name}",
                'Detail_Type': 'Employee_Header'
            }
            
            # Create data rows for this employee
            attendance_data = []
            in_time_data = []
            out_time_data = []
            status_data = []
            dates_data = []
            
            for day in all_days:
                date_key = day.date()
                date_str = day.strftime('%d-%m-%Y')
                
                # Get times for this date
                in_time = in_times.get(date_key)
                out_time = out_times.get(date_key)
                
                # Format times
                if pd.notna(in_time):
                    in_time_str = in_time.strftime('%H:%M')
                else:
                    in_time_str = '00:00'
                
                if pd.notna(out_time):
                    out_time_str = out_time.strftime('%H:%M')
                else:
                    out_time_str = '00:00'
                
                # Determine status
                if in_time_str != '00:00' and out_time_str != '00:00':
                    status = 'P'  # Present
                elif in_time_str != '00:00' or out_time_str != '00:00':
                    status = 'E'  # Early departure or incomplete record
                else:
                    status = 'A'  # Absent
                
                # Store data for this date
                dates_data.append(date_str)
                in_time_data.append(in_time_str)
                out_time_data.append(out_time_str)
                status_data.append(status)
            
            # Create rows for this employee
            # Header row
            header_row = {'Employee_Info': f"{emp_id} - {emp_name}", 'Detail_Type': 'Header'}
            for i, date in enumerate(dates_data):
                header_row[f'Day_{i+1:02d}'] = ''
            report_rows.append(header_row)
            
            # In-Time row
            in_row = {'Employee_Info': 'In-Time', 'Detail_Type': 'InTime'}
            for i, in_time in enumerate(in_time_data):
                in_row[f'Day_{i+1:02d}'] = in_time
            report_rows.append(in_row)
            
            # Out-Time row  
            out_row = {'Employee_Info': 'Out-Time', 'Detail_Type': 'OutTime'}
            for i, out_time in enumerate(out_time_data):
                out_row[f'Day_{i+1:02d}'] = out_time
            report_rows.append(out_row)
            
            # Status row
            status_row = {'Employee_Info': 'Status', 'Detail_Type': 'Status'}
            for i, status in enumerate(status_data):
                status_row[f'Day_{i+1:02d}'] = status
            report_rows.append(status_row)
            
            # Date row
            date_row = {'Employee_Info': 'Date', 'Detail_Type': 'Date'}
            for i, date in enumerate(dates_data):
                date_row[f'Day_{i+1:02d}'] = date
            report_rows.append(date_row)
            
            # Add empty row for separation
            empty_row = {'Employee_Info': '', 'Detail_Type': 'Separator'}
            for i in range(len(dates_data)):
                empty_row[f'Day_{i+1:02d}'] = ''
            report_rows.append(empty_row)
        
        # Convert to DataFrame
        report_df = pd.DataFrame(report_rows)
        
        # Reorder columns to have Employee_Info and Detail_Type first, then days in order
        day_columns = [col for col in report_df.columns if col.startswith('Day_')]
        day_columns.sort()
        column_order = ['Employee_Info', 'Detail_Type'] + day_columns
        report_df = report_df[column_order]
        
        return report_df
    
    def generate_summary_report(self, month: str) -> Optional[pd.DataFrame]:
        """Generate a summary report with attendance statistics"""
        month_df = self.df[self.df['YearMonth'] == month].copy()
        
        if month_df.empty:
            return None
        
        # Get date range for the month
        year, month_num = month.split('-')
        start_date = datetime.datetime(int(year), int(month_num), 1)
        
        if int(month_num) == 12:
            end_date = datetime.datetime(int(year) + 1, 1, 1) - datetime.timedelta(days=1)
        else:
            end_date = datetime.datetime(int(year), int(month_num) + 1, 1) - datetime.timedelta(days=1)
        
        total_working_days = (end_date - start_date).days + 1
        
        # Get unique employees
        employees = sorted(month_df[['EmpID', 'EmployeeName']].drop_duplicates().values.tolist())
        
        summary_data = []
        
        for emp_id, emp_name in employees:
            emp_records = month_df[month_df['EmpID'] == emp_id].copy()
            
            # Count unique dates with attendance records
            unique_dates = emp_records['DateTime'].dt.date.nunique()
            present_days = unique_dates
            absent_days = total_working_days - present_days
            attendance_percentage = round((present_days / total_working_days) * 100, 2)
            
            summary_data.append({
                'Employee_ID': emp_id,
                'Employee_Name': emp_name,
                'Total_Working_Days': total_working_days,
                'Present_Days': present_days,
                'Absent_Days': absent_days,
                'Attendance_Percentage': f"{attendance_percentage}%"
            })
        
        return pd.DataFrame(summary_data)
    
    def generate_reports(self, selected_months: List[str]):
        """Generate reports for selected months with enhanced Excel formatting"""
        if not selected_months:
            print("❌ No months selected for report generation")
            return
        
        print(f"\n🎯 Generating reports for {len(selected_months)} months...")
        
        # Create output directory
        output_dir = f"attendance_reports_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate Excel file with multiple sheets
        excel_filename = os.path.join(output_dir, "attendance_report.xlsx")
        
        try:
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                
                for i, month in enumerate(selected_months, 1):
                    print(f"\n📊 Processing month {i}/{len(selected_months)}: {month}")
                    
                    # Generate monthly report
                    monthly_report = self.generate_monthly_report(month)
                    
                    if monthly_report is not None:
                        # Create sheet name
                        year, month_num = month.split('-')
                        sheet_name = datetime.datetime(int(year), int(month_num), 1).strftime("%B-%Y")
                        
                        # Write to Excel (remove index to clean up the display)
                        monthly_report.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"   ✅ Excel sheet created: {sheet_name}")
                        
                        # Also create individual CSV
                        csv_filename = os.path.join(output_dir, f"report_{month}.csv")
                        monthly_report.to_csv(csv_filename, index=False)
                        print(f"   📄 CSV saved: {csv_filename}")
                        
                        # Generate summary report
                        summary_report = self.generate_summary_report(month)
                        if summary_report is not None:
                            summary_sheet_name = f"Summary-{sheet_name}"
                            summary_report.to_excel(writer, sheet_name=summary_sheet_name, index=False)
                            
                            # Save summary CSV
                            summary_csv = os.path.join(output_dir, f"summary_{month}.csv")
                            summary_report.to_csv(summary_csv, index=False)
                            print(f"   📈 Summary saved: summary_{month}.csv")
                    
            print(f"\n🎉 All reports generated successfully!")
            print(f"📁 Reports saved in: {output_dir}")
            print(f"📊 Excel file: {excel_filename}")
            
        except Exception as e:
            print(f"❌ Error generating reports: {str(e)}")
            logging.error(f"Error generating reports: {str(e)}")


def main():
    """Main function with command line argument support"""
    parser = argparse.ArgumentParser(description='Enhanced Attendance Report Generator')
    parser.add_argument('--file', '-f', default='AGL_0001.TXT', 
                       help='Input attendance file (default: AGL_0001.TXT)')
    parser.add_argument('--months', '-m', 
                       help='Comma-separated list of months (YYYY-MM format) or "all"')
    parser.add_argument('--interactive', '-i', action='store_true', 
                       help='Interactive month selection mode')
    
    args = parser.parse_args()
    
    # Initialize the generator
    generator = AttendanceReportGenerator()
    
    # Read and preprocess data
    if not generator.read_attendance_file(args.file):
        return
    
    if not generator.preprocess_data():
        return
    
    # Month selection
    if args.months:
        if args.months.lower() == 'all':
            selected_months = generator.available_months
        else:
            # Parse comma-separated months
            requested_months = [m.strip() for m in args.months.split(',')]
            selected_months = [m for m in requested_months if m in generator.available_months]
            
            if not selected_months:
                print(f"❌ No valid months found in: {args.months}")
                print(f"Available months: {', '.join(generator.available_months)}")
                return
    else:
        # Interactive mode (default)
        selected_months = generator.get_month_selection_menu()
    
    if selected_months:
        generator.generate_reports(selected_months)
    else:
        print("👋 No reports generated. Goodbye!")


if __name__ == "__main__":
    main()