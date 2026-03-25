"""
Command-line interface for the Utilization Report Generator
"""

import sys
import click
import os

from .core import ReportGenerator


@click.command()
@click.argument('input_file', type=click.Path(exists=True), required=False)
@click.option('-o', '--output', 'output_dir', type=click.Path(), help='Output directory for reports')
@click.option('--open', 'auto_open', is_flag=True, help='Auto-open HTML report in browser')
def generate_reports(input_file, output_dir, auto_open):
    """
    Generate QEA-UHG Leave & Utilization Reports.
    
    INPUT_FILE: Path to source Excel file (optional - will prompt if not provided)
    """
    
    # If input file not provided, prompt user
    if not input_file:
        input_file = click.prompt('Enter path to source Excel file')
    
    # Validate input file
    input_file = input_file.strip().strip('"').strip("'")
    if not os.path.exists(input_file):
        click.secho(f"[ERROR] File not found: {input_file}", fg='red')
        sys.exit(1)
    
    try:
        # Generate reports
        generator = ReportGenerator(input_file, output_dir)
        result = generator.generate()
        
        # Open HTML if requested
        if auto_open:
            import webbrowser
            webbrowser.open(f"file:///{result['html_path'].replace(chr(92), '/')}")
            click.secho("\n[INFO] Opening HTML report in browser...", fg='blue')
        
        click.secho("\n✓ Reports generated successfully!", fg='green')
        
    except Exception as e:
        click.secho(f"\n[ERROR] {str(e)}", fg='red')
        sys.exit(1)


@click.group()
def cli():
    """QEA-UHG Leave & Utilization Report Generator"""
    pass


cli.add_command(generate_reports, 'generate')


if __name__ == '__main__':
    cli()
