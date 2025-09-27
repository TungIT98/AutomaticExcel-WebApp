#!/usr/bin/env python3
"""
Script deploy lên Oracle Cloud
Chạy trên server Ubuntu
"""

import os
import subprocess
import sys

def install_dependencies():
    """Cài đặt dependencies"""
    print("🔧 Cài đặt dependencies...")
    
    # Update system
    subprocess.run(["sudo", "apt", "update"], check=True)
    subprocess.run(["sudo", "apt", "upgrade", "-y"], check=True)
    
    # Install Python and pip
    subprocess.run(["sudo", "apt", "install", "-y", "python3", "python3-pip", "python3-venv"], check=True)
    
    # Install system dependencies
    subprocess.run(["sudo", "apt", "install", "-y", "gcc", "g++", "libffi-dev", "libssl-dev"], check=True)
    
    print("✅ Dependencies installed!")

def setup_application():
    """Setup application"""
    print("📁 Setup application...")
    
    # Create app directory
    os.makedirs("/opt/excel-word-converter", exist_ok=True)
    os.chdir("/opt/excel-word-converter")
    
    # Create virtual environment
    subprocess.run(["python3", "-m", "venv", "venv"], check=True)
    
    # Activate virtual environment
    activate_script = "/opt/excel-word-converter/venv/bin/activate"
    
    # Install Python packages
    subprocess.run([
        "bash", "-c", 
        f"source {activate_script} && pip install -r requirements.txt"
    ], check=True)
    
    print("✅ Application setup complete!")

def setup_nginx():
    """Setup Nginx reverse proxy"""
    print("🌐 Setup Nginx...")
    
    # Install Nginx
    subprocess.run(["sudo", "apt", "install", "-y", "nginx"], check=True)
    
    # Create Nginx config
    nginx_config = """
server {
    listen 80;
    server_name _;
    
    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
"""
    
    with open("/etc/nginx/sites-available/excel-word-converter", "w") as f:
        f.write(nginx_config)
    
    # Enable site
    subprocess.run(["sudo", "ln", "-sf", "/etc/nginx/sites-available/excel-word-converter", "/etc/nginx/sites-enabled/"], check=True)
    subprocess.run(["sudo", "rm", "-f", "/etc/nginx/sites-enabled/default"], check=True)
    subprocess.run(["sudo", "nginx", "-t"], check=True)
    subprocess.run(["sudo", "systemctl", "restart", "nginx"], check=True)
    
    print("✅ Nginx setup complete!")

def setup_systemd():
    """Setup systemd service"""
    print("🔧 Setup systemd service...")
    
    service_config = """
[Unit]
Description=Excel to Word Converter
After=network.target

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/opt/excel-word-converter
Environment=PATH=/opt/excel-word-converter/venv/bin
ExecStart=/opt/excel-word-converter/venv/bin/python app.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
"""
    
    with open("/etc/systemd/system/excel-word-converter.service", "w") as f:
        f.write(service_config)
    
    # Enable and start service
    subprocess.run(["sudo", "systemctl", "daemon-reload"], check=True)
    subprocess.run(["sudo", "systemctl", "enable", "excel-word-converter"], check=True)
    subprocess.run(["sudo", "systemctl", "start", "excel-word-converter"], check=True)
    
    print("✅ Systemd service setup complete!")

def main():
    """Main deployment function"""
    print("🚀 Starting Oracle Cloud deployment...")
    
    try:
        install_dependencies()
        setup_application()
        setup_nginx()
        setup_systemd()
        
        print("🎉 Deployment complete!")
        print("📱 Your app is running at: http://YOUR_SERVER_IP")
        print("🔧 To check status: sudo systemctl status excel-word-converter")
        print("📋 To view logs: sudo journalctl -u excel-word-converter -f")
        
    except Exception as e:
        print(f"❌ Deployment failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
