from flask import Flask, request, jsonify
import os

app = Flask(__name__)

@app.route('/')
def home():
    return jsonify({
        'message': 'Automatic Excel to Word Converter',
        'status': 'running',
        'endpoints': {
            'admin': '/admin',
            'api': '/api'
        }
    })

@app.route('/admin')
def admin():
    return jsonify({
        'message': 'Admin Panel',
        'status': 'running'
    })

@app.route('/api/login', methods=['POST'])
def login():
    return jsonify({'message': 'Login endpoint'})

@app.route('/api/register', methods=['POST'])
def register():
    return jsonify({'message': 'Register endpoint'})

# Vercel serverless function
handler = app