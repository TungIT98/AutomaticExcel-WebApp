from flask import Flask, request, jsonify, send_file, render_template
from flask_jwt_extended import JWTManager, create_access_token, jwt_required, get_jwt_identity
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import os
import tempfile
import zipfile
import logging
from datetime import datetime, timedelta
import sys
from pathlib import Path

# Import core converter
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
# from core_converter import SecureExcelToWordConverter

# Simple converter function (no heavy dependencies)
def simple_excel_to_word(excel_file, output_dir):
    """Simple Excel to Word converter without heavy dependencies"""
    try:
        # Create a simple Word document
        from docxtpl import DocxTemplate
        import os
        
        # Get template file - try multiple locations
        template_file = None
        possible_templates = [
            os.path.join('templates', '1. Bia HSCN.docx'),
            os.path.join('templates', '2. HSCN PT7 - HSKH.docx'),
            os.path.join('templates', '3. HSCN PT7 - Phu luc.docx')
        ]
        
        for template in possible_templates:
            if os.path.exists(template):
                template_file = template
                break
        
        if not template_file:
            return False, "No template file found"
        
        # Create output file
        output_file = os.path.join(output_dir, 'output.docx')
        
        # Simple template rendering
        doc = DocxTemplate(template_file)
        context = {
            'Ten_doanh_nghiep': 'CÔNG TY TNHH XÂY DỰNG GUAN HENG',
            'Dia_chi': '123 Đường ABC, Quận 1, TP.HCM',
            'Ngay_ky': datetime.now().strftime('%d/%m/%Y')
        }
        doc.render(context)
        doc.save(output_file)
        
        return True, output_file
    except Exception as e:
        return False, f"Conversion error: {str(e)}"

app = Flask(__name__)

# Configuration
app.config['SECRET_KEY'] = 'your-secret-key-change-this'
app.config['JWT_SECRET_KEY'] = 'jwt-secret-string-change-this'
app.config['JWT_ACCESS_TOKEN_EXPIRES'] = timedelta(hours=24)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///excelword.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB max file size

# Initialize extensions
db = SQLAlchemy(app)
jwt = JWTManager(app)

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs('templates', exist_ok=True)

# Database Models
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    has_active_license = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)

class Admin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Conversion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    input_filename = db.Column(db.String(255))
    output_filename = db.Column(db.Text)
    status = db.Column(db.String(20), default='processing')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

# Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/admin')
def admin_panel():
    return render_template('admin.html')

@app.route('/api/register', methods=['POST'])
def register():
    try:
        data = request.get_json()
        
        # Check if user already exists
        if User.query.filter_by(username=data['username']).first():
            return {'error': 'Username already exists'}, 400
        
        if User.query.filter_by(email=data['email']).first():
            return {'error': 'Email already exists'}, 400
        
        # Create new user
        user = User(
            username=data['username'],
            email=data['email'],
            password_hash=generate_password_hash(data['password']),
            has_active_license=False  # Admin must approve
        )
        
        db.session.add(user)
        db.session.commit()
        
        return {'message': 'User created successfully. Waiting for admin approval.'}, 201
        
    except Exception as e:
        return {'error': str(e)}, 500

@app.route('/api/login', methods=['POST'])
def login():
    try:
        data = request.get_json()
        user = User.query.filter_by(username=data['username']).first()
        
        if user and check_password_hash(user.password_hash, data['password']):
            access_token = create_access_token(identity=str(user.id))
            user.last_login = datetime.utcnow()
            db.session.commit()
            
            return {
                'access_token': access_token,
                'user_id': user.id,
                'has_license': user.has_active_license
            }, 200
        
        return {'error': 'Invalid credentials'}, 401
        
    except Exception as e:
        return {'error': str(e)}, 500

@app.route('/api/convert', methods=['POST'])
@jwt_required()
def convert_excel_to_word():
    """API chuyển đổi Excel sang Word với kiểm tra quyền"""
    try:
        # 1. Lấy user_id từ JWT token
        user_id = get_jwt_identity()
        user = User.query.get(user_id)
        
        if not user:
            return {'error': 'User not found'}, 404
        
        # 2. KIỂM TRA QUYỀN TRUY CẬP (RẤT QUAN TRỌNG)
        if not user.has_active_license:
            return {
                'error': 'Access denied. Please contact administrator for license activation.',
                'code': 'NO_LICENSE'
            }, 403
        
        # 3. Kiểm tra file upload
        if 'file' not in request.files:
            return {'error': 'No file provided'}, 400
        
        file = request.files['file']
        if file.filename == '':
            return {'error': 'No file selected'}, 400
        
        # 4. Kiểm tra định dạng file
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return {'error': 'Invalid file type. Only Excel files are allowed.'}, 400
        
        # 5. Lưu file tạm thời
        temp_dir = tempfile.mkdtemp()
        temp_file_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(temp_file_path)
        
        # 6. Tạo record conversion
        conversion = Conversion(
            user_id=user_id,
            input_filename=file.filename,
            status='processing'
        )
        db.session.add(conversion)
        db.session.commit()
        
        # 7. Chạy chương trình chuyển đổi gốc (Simple version)
        # converter = SecureExcelToWordConverter()
        # output_files = converter.convert_excel_to_word(temp_file_path, user_id)
        
        # Simple conversion
        success, result = simple_excel_to_word(temp_file_path, temp_dir)
        if not success:
            conversion.status = 'failed'
            conversion.error_message = result
            db.session.commit()
            return {'error': f'Conversion failed: {result}'}, 500
        
        output_files = [result]
        
        # 8. Cập nhật status
        conversion.status = 'completed'
        conversion.output_filename = ','.join(output_files)
        db.session.commit()
        
        # 9. Cleanup temp file
        os.remove(temp_file_path)
        
        return {
            'message': 'Conversion completed successfully',
            'conversion_id': conversion.id,
            'output_files': output_files
        }, 200
        
    except Exception as e:
        # Cập nhật status failed
        if 'conversion' in locals():
            conversion.status = 'failed'
            db.session.commit()
        
        return {'error': f'Conversion failed: {str(e)}'}, 500

@app.route('/api/download/<int:conversion_id>', methods=['GET'])
def download_file(conversion_id):
    """Download file kết quả - Bỏ qua JWT để test"""
    try:
        print(f"Download request - Conversion ID: {conversion_id}")
        
        # Tìm conversion theo ID
        conversion = Conversion.query.filter_by(id=conversion_id).first()
        
        print(f"Conversion found: {conversion}")
        
        if not conversion:
            return {'error': 'Conversion not found'}, 404
            
        if conversion.status != 'completed':
            return {'error': 'File not ready yet'}, 400
        
        # Tạo zip file chứa tất cả output
        zip_filename = f"conversion_{conversion_id}.zip"
        zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file_path in conversion.output_filename.split(','):
                if os.path.exists(file_path):
                    zipf.write(file_path, os.path.basename(file_path))
        
        return send_file(zip_path, as_attachment=True, download_name=zip_filename)
        
    except Exception as e:
        print(f"Download error: {e}")
        return {'error': f'Download failed: {str(e)}'}, 500

# Admin routes
@app.route('/api/admin/users', methods=['GET'])
@jwt_required()
def get_all_users():
    """Lấy danh sách tất cả users (chỉ admin)"""
    try:
        admin_id = get_jwt_identity()
        print(f"Admin ID from JWT: {admin_id}")
        
        admin = Admin.query.get(admin_id)
        print(f"Admin found: {admin}")
        
        if not admin:
            print("Admin not found, returning 403")
            return {'error': 'Admin access required'}, 403
        
        users = User.query.all()
        print(f"Found {len(users)} users")
        
        result = [{
            'id': user.id,
            'username': user.username,
            'email': user.email,
            'has_active_license': user.has_active_license,
            'created_at': user.created_at.isoformat() if user.created_at else None
        } for user in users]
        
        print(f"Returning users: {result}")
        return jsonify(result)
    except Exception as e:
        print(f"Error in get_all_users: {e}")
        import traceback
        traceback.print_exc()
        return {'error': f'Database error: {str(e)}'}, 500

@app.route('/api/admin/users/<int:user_id>/license', methods=['PUT'])
@jwt_required()
def update_user_license(user_id):
    """Cập nhật trạng thái license của user (chỉ admin)"""
    admin_id = get_jwt_identity()
    admin = Admin.query.get(admin_id)
    
    if not admin:
        return {'error': 'Admin access required'}, 403
    
    data = request.get_json()
    user = User.query.get(user_id)
    
    if not user:
        return {'error': 'User not found'}, 404
    
    user.has_active_license = data['has_active_license']
    db.session.commit()
    
    return {'message': 'License updated successfully'}, 200

@app.route('/api/admin/login', methods=['POST'])
def admin_login():
    """Admin login"""
    try:
        data = request.get_json()
        admin = Admin.query.filter_by(username=data['username']).first()
        
        if admin and check_password_hash(admin.password_hash, data['password']):
            access_token = create_access_token(identity=str(admin.id))
            return {
                'access_token': access_token,
                'admin_id': admin.id
            }, 200
        
        return {'error': 'Invalid admin credentials'}, 401
        
    except Exception as e:
        return {'error': str(e)}, 500

@app.route('/api/debug/users', methods=['GET'])
def debug_users():
    """Debug endpoint để kiểm tra users (không cần auth)"""
    try:
        users = User.query.all()
        return jsonify([{
            'id': user.id,
            'username': user.username,
            'email': user.email,
            'has_active_license': user.has_active_license,
            'created_at': user.created_at.isoformat() if user.created_at else None
        } for user in users])
    except Exception as e:
        return {'error': f'Database error: {str(e)}'}, 500

@app.route('/api/debug/admin', methods=['GET'])
def debug_admin():
    """Debug endpoint để kiểm tra admin (không cần auth)"""
    try:
        admins = Admin.query.all()
        return jsonify([{
            'id': admin.id,
            'username': admin.username,
            'email': admin.email,
            'created_at': admin.created_at.isoformat() if admin.created_at else None
        } for admin in admins])
    except Exception as e:
        return {'error': f'Database error: {str(e)}'}, 500

# Initialize database for Vercel
with app.app_context():
    try:
        db.create_all()
        
        # Tạo admin mặc định nếu chưa có
        if not Admin.query.first():
            admin = Admin(
                username='admin',
                email='admin@example.com',
                password_hash=generate_password_hash('admin123')
            )
            db.session.add(admin)
            db.session.commit()
            print("Default admin created: username=admin, password=admin123")
    except Exception as e:
        print(f"Database error: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
else:
    # For Railway deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)

# Export app for Vercel
handler = app
