#!/usr/bin/env bash
set -euo pipefail

# ===== 配置区（如需可改） =====
REPO_URL="https://github.com/asliujinhe/TeachingScriptGenerator"
APP_USER="teachingapp"
APP_DIR="/opt/TeachingScriptGenerator"
PYTHON_BIN="python3"              # Ubuntu22.04 自带 python3
PORT="9003"
SERVICE_NAME="teachingscript.service"
WORKERS="3"
TIMEOUT="120"                     # gunicorn 超时时间
UFW_OPEN_PORT="true"              # 系统启用 ufw 时，是否自动放行 9003 端口

echo "==> 检查/安装基础依赖..."
export DEBIAN_FRONTEND=noninteractive
apt-get update -y
apt-get install -y git ${PYTHON_BIN} python3-venv python3-pip

# 创建系统用户（不可登录）
if ! id -u "${APP_USER}" >/dev/null 2>&1; then
  echo "==> 创建系统用户 ${APP_USER}"
  useradd -r -s /usr/sbin/nologin -d "${APP_DIR}" "${APP_USER}"
fi

# 目录准备
mkdir -p "${APP_DIR}"
chown -R "${APP_USER}:${APP_USER}" "${APP_DIR}"

# 克隆/更新仓库
if [ ! -d "${APP_DIR}/.git" ]; then
  echo "==> 克隆仓库 ${REPO_URL} 到 ${APP_DIR}"
  git clone --depth 1 "${REPO_URL}" "${APP_DIR}"
  chown -R "${APP_USER}:${APP_USER}" "${APP_DIR}"
else
  echo "==> 仓库已存在，执行 git pull"
  pushd "${APP_DIR}" >/dev/null
  sudo -u "${APP_USER}" git pull --rebase --autostash
  popd >/dev/null
fi

# Python 虚拟环境与依赖
echo "==> 创建/更新虚拟环境并安装依赖"
if [ ! -d "${APP_DIR}/venv" ]; then
  sudo -u "${APP_USER}" ${PYTHON_BIN} -m venv "${APP_DIR}/venv"
fi
# 升级 pip
sudo -u "${APP_USER}" "${APP_DIR}/venv/bin/pip" install --upgrade pip wheel setuptools
# 安装 requirements.txt
if [ -f "${APP_DIR}/requirements.txt" ]; then
  sudo -u "${APP_USER}" "${APP_DIR}/venv/bin/pip" install -r "${APP_DIR}/requirements.txt"
else
  echo "!! 警告：未发现 ${APP_DIR}/requirements.txt，跳过依赖安装"
fi
# 确保 gunicorn 存在（有些 requirements 里未列出）
sudo -u "${APP_USER}" "${APP_DIR}/venv/bin/pip" install gunicorn

# systemd 服务单元
UNIT_FILE="/etc/systemd/system/${SERVICE_NAME}"
echo "==> 生成 systemd 单元：${UNIT_FILE}"
cat > "${UNIT_FILE}" <<EOF
[Unit]
Description=TeachingScriptGenerator Gunicorn Service
After=network.target

[Service]
Type=simple
User=${APP_USER}
Group=${APP_USER}
WorkingDirectory=${APP_DIR}
Environment=PYTHONUNBUFFERED=1
# 如需传参给应用，可在此处添加：Environment=PORT=${PORT}
ExecStart=${APP_DIR}/venv/bin/gunicorn \\
  --workers ${WORKERS} \\
  --timeout ${TIMEOUT} \\
  --bind 0.0.0.0:${PORT} \\
  app:app
Restart=always
RestartSec=3

# 提高文件描述符上限（可选）
LimitNOFILE=65535

[Install]
WantedBy=multi-user.target
EOF

# 权限与加载
echo "==> 重新加载 systemd，启用并启动服务"
systemctl daemon-reload
systemctl enable "${SERVICE_NAME}"
systemctl restart "${SERVICE_NAME}"

# 防火墙（可选）
if command -v ufw >/dev/null 2>&1; then
  if ufw status | grep -q "Status: active"; then
    if [ "${UFW_OPEN_PORT}" = "true" ]; then
      echo "==> UFW 已启用，放行端口 ${PORT}"
      ufw allow "${PORT}"/tcp || true
    fi
  fi
fi

echo "==> 部署完成"
echo "服务名: ${SERVICE_NAME}"
echo "监听端口: ${PORT}"
echo "目录: ${APP_DIR}"
echo
echo "常用命令："
echo "  journalctl -u ${SERVICE_NAME} -f    # 实时日志"
echo "  systemctl status ${SERVICE_NAME}     # 查看状态"
echo "  systemctl restart ${SERVICE_NAME}    # 重启服务"
echo
echo "访问：http://<服务器IP>:${PORT}/"
