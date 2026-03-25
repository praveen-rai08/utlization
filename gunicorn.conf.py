import os

bind = f"0.0.0.0:{os.environ.get('PORT', '5000')}"
workers = 2
threads = 4
timeout = 120          # allow up to 2 min for large file processing
worker_class = "sync"
accesslog = "-"        # log to stdout
errorlog = "-"
loglevel = "info"
