# gunicorn.conf.py
timeout = 120  # segundos
worker_class = 'sync'
workers = 1  # ya que la memoria es limitada