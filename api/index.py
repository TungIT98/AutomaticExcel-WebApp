# Import the main Flask app
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from app import app
    handler = app
except Exception as e:
    # Fallback simple app
    from flask import Flask
    app = Flask(__name__)
    
    @app.route('/')
    def home():
        return {'message': 'App is running', 'error': str(e)}
    
    handler = app