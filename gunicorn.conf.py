# gunicorn.conf.py
timeout = 600  # 10 минут таймаут
keepalive = 5
workers = 2
worker_class = 'sync'
max_requests = 1000
max_requests_jitter = 100
