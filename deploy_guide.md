# 🚀 Hướng dẫn Deploy lên Oracle Cloud

## 📋 Bước 1: Tạo tài khoản Oracle Cloud
1. Truy cập: https://cloud.oracle.com/
2. Đăng ký tài khoản miễn phí
3. Xác thực thẻ tín dụng (không bị trừ tiền)
4. Chọn region gần nhất (Singapore, Tokyo)

## 📋 Bước 2: Tạo VM Instance
1. Vào Console → Compute → Instances
2. Create Instance
3. Chọn: Always Free Eligible
4. OS: Ubuntu 20.04 LTS
5. Shape: VM.Standard.E2.1.Micro (1GB RAM, 1 CPU)
6. Storage: 10GB (miễn phí)
7. SSH Key: Tạo key mới hoặc upload key hiện có

## 📋 Bước 3: Kết nối SSH
```bash
ssh -i your-key.pem ubuntu@YOUR_SERVER_IP
```

## 📋 Bước 4: Upload code lên server
```bash
# Tạo thư mục
mkdir -p /opt/excel-word-converter

# Upload code (sử dụng SCP hoặc Git)
scp -i your-key.pem -r webapp/* ubuntu@YOUR_SERVER_IP:/opt/excel-word-converter/
```

## 📋 Bước 5: Chạy script deploy
```bash
cd /opt/excel-word-converter
python3 deploy_oracle.py
```

## 📋 Bước 6: Mở port firewall
```bash
# Mở port 80 và 443
sudo ufw allow 80
sudo ufw allow 443
sudo ufw enable
```

## 📋 Bước 7: Kiểm tra
- Truy cập: http://YOUR_SERVER_IP
- Admin Panel: http://YOUR_SERVER_IP/admin
- Admin: admin / admin123

## 🔧 Troubleshooting
```bash
# Kiểm tra service
sudo systemctl status excel-word-converter

# Xem logs
sudo journalctl -u excel-word-converter -f

# Restart service
sudo systemctl restart excel-word-converter

# Kiểm tra Nginx
sudo nginx -t
sudo systemctl status nginx
```

## 💰 Chi phí: HOÀN TOÀN MIỄN PHÍ
- VM: Miễn phí vĩnh viễn
- Storage: 10GB miễn phí
- Bandwidth: Không giới hạn
- Không có thời hạn
