# Simple Flask app for Vercel
from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/')
def home():
    return jsonify({
        'message': 'Automatic Excel to Word Converter',
        'status': 'running',
        'version': '1.0.0'
    })

@app.route('/api/health')
def health():
    return jsonify({'status': 'ok'})

# Vercel handler
handler = app