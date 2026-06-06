#!/bin/bash

set -e

echo "=== VBA到WPS JS宏转换器部署脚本 ==="

APP_DIR="/opt/vba-converter"
VENV_DIR="$APP_DIR/venv"
PYTHON_VERSION="3.10"

echo "1. 检查系统环境..."

if ! command -v python$PYTHON_VERSION &> /dev/null; then
    echo "安装 Python $PYTHON_VERSION..."
    sudo apt update && sudo apt install -y python$PYTHON_VERSION python$PYTHON_VERSION-venv python$PYTHON_VERSION-dev
fi

if ! command -v node &> /dev/null; then
    echo "安装 Node.js..."
    curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
    sudo apt install -y nodejs
fi

echo "2. 创建应用目录..."
sudo mkdir -p $APP_DIR
sudo chown $USER:$USER $APP_DIR

echo "3. 复制代码文件..."
cp -r ./* $APP_DIR/

echo "4. 创建 Python 虚拟环境..."
python$PYTHON_VERSION -m venv $VENV_DIR
source $VENV_DIR/bin/activate

echo "5. 安装 Python 依赖..."
pip install --upgrade pip
pip install -r $APP_DIR/requirements.txt

echo "6. 安装前端依赖..."
cd $APP_DIR/frontend
npm install
npm run build

echo "7. 创建数据目录..."
mkdir -p $APP_DIR/data/uploads $APP_DIR/data/converted $APP_DIR/data/logs

echo "8. 创建系统服务..."
cat > /tmp/vba-converter.service << EOF
[Unit]
Description=VBA to WPS JS Macro Converter Service
After=network.target

[Service]
User=$USER
WorkingDirectory=$APP_DIR
Environment="FLASK_ENV=production"
Environment="PATH=$VENV_DIR/bin:$PATH"
ExecStart=$VENV_DIR/bin/gunicorn --workers=4 --bind=0.0.0.0:5000 run:app
Restart=always

[Install]
WantedBy=multi-user.target
EOF

sudo mv /tmp/vba-converter.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable vba-converter
sudo systemctl start vba-converter

echo "9. 配置 Nginx 反向代理（可选）..."
read -p "是否配置 Nginx 反向代理? (y/n): " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    if ! command -v nginx &> /dev/null; then
        sudo apt install -y nginx
    fi
    
    read -p "输入域名（如 converter.example.com）: " DOMAIN
    
    cat > /tmp/vba-converter.conf << EOF
server {
    listen 80;
    server_name $DOMAIN;

    location / {
        root $APP_DIR/frontend/dist;
        try_files \$uri \$uri/ /index.html;
    }

    location /api/ {
        proxy_pass http://localhost:5000/api/;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
    }

    location /health {
        proxy_pass http://localhost:5000/health;
    }
}
EOF

    sudo mv /tmp/vba-converter.conf /etc/nginx/sites-available/vba-converter
    sudo ln -sf /etc/nginx/sites-available/vba-converter /etc/nginx/sites-enabled/
    sudo nginx -t
    sudo systemctl reload nginx
    
    echo "建议: 运行 'sudo certbot --nginx -d $DOMAIN' 配置 HTTPS"
fi

echo ""
echo "=== 部署完成 ==="
echo ""
echo "服务已启动，运行在 http://localhost:5000"
echo ""
echo "配置 LLM API:"
echo "1. 访问 http://localhost:5000"
echo "2. 在 LLM配置 部分输入您的 API Key、端点和模型名称"
echo "3. 点击保存配置"
echo ""
echo "服务管理命令:"
echo "  启动: sudo systemctl start vba-converter"
echo "  停止: sudo systemctl stop vba-converter"
echo "  重启: sudo systemctl restart vba-converter"
echo "  状态: sudo systemctl status vba-converter"
echo "  日志: sudo journalctl -u vba-converter -f"