# wsgi.py — 生产环境入口（给 gunicorn 用）
# 假设你的 Flask 实例在仓库根目录的 app.py 里，变量名为 app
from app import app as application
# 同时暴露名为 app 的变量，方便用 "wsgi:app" 这种写法
app = application
