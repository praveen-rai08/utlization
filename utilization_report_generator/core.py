"""
Main ReportGenerator class that orchestrates the report generation pipeline
"""

import os
from datetime import date

from .reader import read_source_data
from .generator import generate_excel_report, generate_html_dashboard


class ReportGenerator:
    """Main class for generating utilization reports"""
    
    def __init__(self, source_file, output_dir=None):
        """
        Initialize the report generator.
        
        Args:
            source_file (str): Path to source Excel file
            output_dir (str): Output directory (defaults to same as source file)
        """
        self.source_file = source_file
        self.output_dir = output_dir or os.path.dirname(os.path.abspath(source_file))
        self.records = None
        self.months_order = None
        self.stats = None
    
    def generate(self):
        """
        Generate both Excel and HTML reports.
        
        Returns:
            dict: Contains paths to generated files and summary statistics
        """
        print("=" * 60)
        print("  QEA – UHG Leave & Utilization Report Generator")
        print("=" * 60)
        
        # Step 1: Read source data
        self.records, self.months_order, self.stats = read_source_data(self.source_file)
        
        # Step 2: Generate outputs
        today_str = date.today().strftime("%Y-%m-%d")
        xlsx_out = os.path.join(self.output_dir, f"QEA-UHG-Utilization-Report-{today_str}.xlsx")
        html_out = os.path.join(self.output_dir, f"QEA-UHG-Utilization-Dashboard-{today_str}.html")
        
        # Generate Excel report
        avg_utils, overall_avg, low_count, med_count, high_count, tot_forecast, tot_actual = \
            generate_excel_report(self.records, self.months_order, xlsx_out)
        
        # Generate HTML dashboard
        generate_html_dashboard(self.records, self.months_order, html_out,
                               overall_avg, low_count, med_count, high_count,
                               tot_forecast, tot_actual)
        
        print(f"\n[4/4] Done!")
        print(f"\n  Excel  : {xlsx_out}")
        print(f"  HTML   : {html_out}")
        print(f"\n  Associates : {len(self.records)}")
        print(f"  Avg H1 Util: {overall_avg}%")
        print(f"  High (>={UTIL_HIGH}%): {high_count}  |  Medium ({UTIL_MEDIUM}-{UTIL_HIGH-1}%): {med_count}  |  Low (<{UTIL_MEDIUM}%): {low_count}")
        
        return {
            'excel_path': xlsx_out,
            'html_path': html_out,
            'total_associates': len(self.records),
            'overall_avg_util': overall_avg,
            'high_util_count': high_count,
            'medium_util_count': med_count,
            'low_util_count': low_count,
            'total_forecast': tot_forecast,
            'total_actual': tot_actual,
        }


# Import after class definition to avoid circular imports
from .config import UTIL_HIGH, UTIL_MEDIUM
