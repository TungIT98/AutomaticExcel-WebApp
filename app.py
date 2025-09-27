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

# Import core converter from parent directory
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from excel_reader import ExcelReader

# Core converter function - replicates run.py logic with robust error handling
def core_excel_to_word(excel_file, output_dir):
    """Core Excel to Word converter - replicates run.py logic with robust error handling"""
    try:
        from docxtpl import DocxTemplate
        from pathlib import Path
        import os
        import pandas as pd
        
        print(f"=== BẮT ĐẦU CHUYỂN ĐỔI ===")
        print(f"Excel file: {excel_file}")
        print(f"Output dir: {output_dir}")
        
        # 1. Validate Excel file first
        try:
            print("Validating Excel file...")
            # Try to read Excel file to check if it's valid
            test_df = pd.read_excel(excel_file, sheet_name=None, nrows=1)
            print(f"Excel file validation successful. Found sheets: {list(test_df.keys())}")
        except Exception as e:
            return False, f"Invalid Excel file: {str(e)}"
        
        # 2. Get template folder - check both parent and current directory
        template_folder = os.path.join('..', 'templates')
        template_folder_path = Path(template_folder)
        
        # If parent templates not found, try current directory
        if not template_folder_path.exists():
            template_folder = 'templates'
            template_folder_path = Path(template_folder)
            print(f"Trying current directory templates: {template_folder_path}")
        
        # If still not found, try absolute path
        if not template_folder_path.exists():
            template_folder = os.path.join(os.getcwd(), 'templates')
            template_folder_path = Path(template_folder)
            print(f"Trying absolute path templates: {template_folder_path}")
        
        if not template_folder_path.exists():
            return False, f"Template folder not found: {template_folder}"
        
        # 3. Find all template files
        print(f"DEBUG: Looking for templates in: {template_folder_path}")
        print(f"DEBUG: Template folder exists: {template_folder_path.exists()}")
        
        templates = list(template_folder_path.glob("*.docx"))
        print(f"DEBUG: Found {len(templates)} .docx files before filtering")
        
        # Filter out temp files and problematic templates
        templates = [t for t in templates if not t.name.startswith('~$') and t.name != '9. Phieu YCTN TNN.docx']
        
        print(f"Found {len(templates)} templates: {[t.name for t in templates]}")
        print(f"DEBUG: Template paths: {[str(t) for t in templates]}")
        
        if not templates:
            return False, "No valid templates found"
        
        # 4. Initialize ExcelReader with error handling
        try:
            print("Initializing ExcelReader...")
            excel_reader = ExcelReader(excel_file)
            print("ExcelReader initialized successfully")
        except Exception as e:
            return False, f"Failed to read Excel file: {str(e)}"
        
        # 5. Get all hoso codes with validation
        try:
            hoso_codes = excel_reader.get_all_hoso_codes()
            print(f"Found {len(hoso_codes)} hoso codes: {hoso_codes}")
            
            if not hoso_codes:
                return False, "No hồ sơ codes found in Excel file. Please check your Excel file has the correct format."
        except Exception as e:
            return False, f"Failed to extract hồ sơ codes: {str(e)}"
        
        output_files = []
        error_count = 0
        
        # 6. Process each hoso code
        for ma_ho_so in hoso_codes:
            print(f"--- Processing hoso: {ma_ho_so} ---")
            
            try:
                # Create context for this hoso
                context = excel_reader.create_complete_context(ma_ho_so)
                print(f"Context keys: {list(context.keys())}")
                
                if not context:
                    print(f"Warning: Empty context for hoso {ma_ho_so}")
                    continue
                
            except Exception as e:
                print(f"Error creating context for hoso {ma_ho_so}: {e}")
                error_count += 1
                continue
            
            # 7. Process each template
            print(f"DEBUG: About to process {len(templates)} templates for hoso {ma_ho_so}")
            for i, template_path in enumerate(templates):
                print(f"DEBUG: Processing template {i+1}/{len(templates)}: {template_path}")
                try:
                    template_filename = template_path.name
                    print(f"  Processing template: {template_filename}")
                    
                    # Create final context
                    final_context = context.copy()
                    
                    # Load and render template
                    doc = DocxTemplate(template_path)
                    doc.render(final_context)
                    
                    # Create output filename
                    ten_kh = context.get('ten_kh', 'Unknown').replace(' ', '_').replace('/', '_').replace('\\', '_')
                    ma_ho_so_clean = ma_ho_so.replace('/', '_').replace('\\', '_')
                    template_name_clean = template_path.stem.replace(' ', '_').replace('/', '_').replace('\\', '_')
                    output_filename = f"{template_name_clean}_{ma_ho_so_clean}_{ten_kh}.docx"
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # Save the document
                    doc.save(output_path)
                    output_files.append(output_path)
                    print(f"    -> Created: {output_path}")
                    
                except Exception as e:
                    print(f"    -> ERROR processing template '{template_filename}': {e}")
                    error_count += 1
                    continue
        
        print(f"=== COMPLETED: Created {len(output_files)} files, {error_count} errors ===")
        
        if not output_files:
            return False, "No files were created. Please check your Excel file format and templates."
        
        return True, output_files
        
    except Exception as e:
        print(f"=== CRITICAL ERROR in core_excel_to_word: {e} ===")
        import traceback
        traceback.print_exc()
        return False, f"Critical error: {str(e)}"

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
        print(f"DEBUG: JWT user_id: {user_id}")
        print(f"DEBUG: JWT user_id type: {type(user_id)}")
        
        # Convert to int if string
        if isinstance(user_id, str):
            user_id = int(user_id)
            print(f"DEBUG: Converted user_id to int: {user_id}")
        
        user = User.query.get(user_id)
        print(f"DEBUG: User found: {user}")
        
        if not user:
            print("DEBUG: User not found in database")
            return {'error': 'User not found'}, 404
        
        # 2. KIỂM TRA QUYỀN TRUY CẬP (RẤT QUAN TRỌNG)
        print(f"DEBUG: User license status: {user.has_active_license}")
        if not user.has_active_license:
            print("DEBUG: User license not active")
            return {
                'error': 'Access denied. Please contact administrator for license activation.',
                'code': 'NO_LICENSE'
            }, 403
        
        # 3. Kiểm tra file upload với validation chi tiết
        print(f"DEBUG: request.files keys: {list(request.files.keys())}")
        print(f"DEBUG: request.files: {request.files}")
        print(f"DEBUG: request.content_type: {request.content_type}")
        print(f"DEBUG: request.content_length: {request.content_length}")
        
        if 'file' not in request.files:
            print("DEBUG: No 'file' key in request.files")
            return {'error': 'No file provided. Please select an Excel file.'}, 400
        
        file = request.files['file']
        print(f"DEBUG: File object: {file}")
        print(f"DEBUG: File filename: {file.filename}")
        print(f"DEBUG: File content_type: {file.content_type}")
        
        if file.filename == '' or file.filename is None:
            print("DEBUG: Empty filename")
            return {'error': 'No file selected. Please choose an Excel file.'}, 400
        
        # 4. Kiểm tra định dạng file với validation chi tiết
        filename_lower = file.filename.lower()
        if not (filename_lower.endswith('.xlsx') or filename_lower.endswith('.xls')):
            return {
                'error': 'Invalid file type. Only Excel files (.xlsx, .xls) are allowed.',
                'received_type': file.content_type,
                'filename': file.filename
            }, 400
        
        # 5. Kiểm tra kích thước file
        file.seek(0, 2)  # Seek to end
        file_size = file.tell()
        file.seek(0)  # Reset to beginning
        
        print(f"DEBUG: File size: {file_size} bytes")
        
        if file_size == 0:
            return {'error': 'File is empty. Please upload a valid Excel file.'}, 400
        
        if file_size > app.config['MAX_CONTENT_LENGTH']:
            return {
                'error': f'File too large. Maximum size allowed: {app.config["MAX_CONTENT_LENGTH"] / (1024*1024):.1f} MB',
                'file_size': f'{file_size / (1024*1024):.1f} MB'
            }, 400
        
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
        
        # 7. Chạy chương trình chuyển đổi gốc (Core version) với error handling
        try:
            print(f"Starting conversion for file: {temp_file_path}")
            success, result = core_excel_to_word(temp_file_path, temp_dir)
            
            if not success:
                print(f"Conversion failed: {result}")
                conversion.status = 'failed'
                conversion.error_message = str(result)
                db.session.commit()
                
                # Clean up temp file
                try:
                    os.remove(temp_file_path)
                except:
                    pass
                
                return {
                    'error': f'Conversion failed: {result}',
                    'details': 'Please check your Excel file format and try again.'
                }, 500
            
            # result is now a list of output files
            output_files = result if isinstance(result, list) else [result]
            print(f"Conversion successful. Created {len(output_files)} files")
            
        except Exception as e:
            print(f"Critical error during conversion: {e}")
            conversion.status = 'failed'
            conversion.error_message = f'Critical error: {str(e)}'
            db.session.commit()
            
            # Clean up temp file
            try:
                os.remove(temp_file_path)
            except:
                pass
            
            return {
                'error': f'Critical conversion error: {str(e)}',
                'details': 'Please contact support if this error persists.'
            }, 500
        
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

@app.route('/api/debug/database', methods=['GET'])
def debug_database():
    """Debug endpoint để kiểm tra toàn bộ database"""
    try:
        users = User.query.all()
        admins = Admin.query.all()
        conversions = Conversion.query.all()
        
        return jsonify({
            'users': [{
                'id': user.id,
                'username': user.username,
                'email': user.email,
                'has_active_license': user.has_active_license,
                'created_at': user.created_at.isoformat() if user.created_at else None
            } for user in users],
            'admins': [{
                'id': admin.id,
                'username': admin.username,
                'email': admin.email,
                'created_at': admin.created_at.isoformat() if admin.created_at else None
            } for admin in admins],
            'conversions': [{
                'id': conv.id,
                'user_id': conv.user_id,
                'status': conv.status,
                'created_at': conv.created_at.isoformat() if conv.created_at else None
            } for conv in conversions],
            'total_users': len(users),
            'total_admins': len(admins),
            'total_conversions': len(conversions)
        })
    except Exception as e:
        return {'error': f'Database error: {str(e)}'}, 500

@app.route('/api/debug/activate-user/<int:user_id>', methods=['POST'])
def activate_user_debug(user_id):
    """Debug endpoint để cấp quyền cho user (không cần auth)"""
    try:
        user = User.query.get(user_id)
        if not user:
            return {'error': 'User not found'}, 404
        
        user.has_active_license = True
        db.session.commit()
        
        return {
            'message': f'User {user.username} has been activated',
            'user_id': user.id,
            'username': user.username,
            'has_active_license': user.has_active_license
        }
    except Exception as e:
        return {'error': f'Database error: {str(e)}'}, 500

@app.route('/api/validate-excel', methods=['POST'])
@jwt_required()
def validate_excel():
    """Validate Excel file before conversion"""
    try:
        user_id = get_jwt_identity()
        user = User.query.get(user_id)
        
        if not user:
            return {'error': 'User not found'}, 404
        
        if not user.has_active_license:
            return {
                'error': 'Access denied. Please contact administrator for license activation.',
                'code': 'NO_LICENSE'
            }, 403
        
        # Check file upload
        if 'file' not in request.files:
            return {'error': 'No file provided'}, 400
        
        file = request.files['file']
        if file.filename == '' or file.filename is None:
            return {'error': 'No file selected'}, 400
        
        # Check file type
        filename_lower = file.filename.lower()
        if not (filename_lower.endswith('.xlsx') or filename_lower.endswith('.xls')):
            return {'error': 'Invalid file type. Only Excel files (.xlsx, .xls) are allowed.'}, 400
        
        # Check file size
        file.seek(0, 2)
        file_size = file.tell()
        file.seek(0)
        
        if file_size == 0:
            return {'error': 'File is empty'}, 400
        
        if file_size > app.config['MAX_CONTENT_LENGTH']:
            return {
                'error': f'File too large. Maximum size: {app.config["MAX_CONTENT_LENGTH"] / (1024*1024):.1f} MB'
            }, 400
        
        # Save temp file for validation
        temp_dir = tempfile.mkdtemp()
        temp_file_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(temp_file_path)
        
        try:
            # Validate Excel file structure
            import pandas as pd
            excel_data = pd.read_excel(temp_file_path, sheet_name=None, nrows=1)
            
            # Check for required sheets
            required_sheets = ['hoso', 'hanghoa']
            missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_data]
            
            if missing_sheets:
                return {
                    'error': f'Missing required sheets: {missing_sheets}',
                    'details': 'Excel file must contain "hoso" and "hanghoa" sheets.'
                }, 400
            
            # Try to initialize ExcelReader
            excel_reader = ExcelReader(temp_file_path)
            hoso_codes = excel_reader.get_all_hoso_codes()
            
            if not hoso_codes:
                return {
                    'error': 'No hồ sơ codes found',
                    'details': 'Excel file must contain valid hồ sơ data.'
                }, 400
            
            # Clean up
            os.remove(temp_file_path)
            
            return {
                'valid': True,
                'sheets': list(excel_data.keys()),
                'hoso_count': len(hoso_codes),
                'hoso_codes': hoso_codes[:5],  # Show first 5 codes
                'file_size': f'{file_size / (1024*1024):.1f} MB'
            }
            
        except Exception as e:
            # Clean up on error
            try:
                os.remove(temp_file_path)
            except:
                pass
            
            return {
                'error': f'Invalid Excel file: {str(e)}',
                'details': 'Please check your Excel file format and structure.'
            }, 400
            
    except Exception as e:
        return {'error': f'Validation error: {str(e)}'}, 500

# Initialize database for Vercel
with app.app_context():
    try:
        db.create_all()
        
        # Tạo admin mặc định nếu chưa có
        existing_admin = Admin.query.filter_by(username='admin').first()
        if not existing_admin:
            try:
                admin = Admin(
                    username='admin',
                    email='admin@example.com',
                    password_hash=generate_password_hash('admin123')
                )
                db.session.add(admin)
                db.session.commit()
                print("Default admin created: username=admin, password=admin123")
            except Exception as e:
                print(f"Admin already exists or error creating admin: {e}")
                db.session.rollback()
        else:
            print("Admin already exists")
            
        # Tự động cấp quyền cho user đầu tiên nếu chưa có license
        first_user = User.query.filter_by(has_active_license=False).first()
        if first_user:
            first_user.has_active_license = True
            db.session.commit()
            print(f"Auto-activated user: {first_user.username}")
    except Exception as e:
        print(f"Database error: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)

# Export app for Vercel
handler = app
