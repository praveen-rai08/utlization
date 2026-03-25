#!/usr/bin/env python
"""
Web UI entry point for Utilization Report Generator
"""

import sys
import os
from utilization_report_generator.web import run_web_app

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='Run Utilization Report Generator Web UI')
    parser.add_argument('-H', '--host', default='127.0.0.1', help='Host to bind to (default: 127.0.0.1)')
    parser.add_argument('-p', '--port', type=int, default=5000, help='Port to bind to (default: 5000)')
    parser.add_argument('-d', '--debug', action='store_true', help='Enable debug mode')
    
    args = parser.parse_args()
    
    print(f"\n{'='*60}")
    print("  QEA – UHG Utilization Report Generator Web UI")
    print(f"{'='*60}")
    print(f"\n  Starting web server on http://{args.host}:{args.port}")
    print(f"  Press Ctrl+C to stop the server\n")
    
    run_web_app(host=args.host, port=args.port, debug=args.debug)
