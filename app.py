from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
import threading
import os
import json
from lasik_bot import LasikBotService  # Import class yang sudah dibuat sebelumnya

app = Flask(__name__)
CORS(app)

# Global variables
bot_instance = None
bot_thread = None
is_bot_running = False
progress_data = {
    'current': 0,
    'total': 0,
    'success': 0,
    'failed': 0,
    'status': 'Idle',
    'is_success': True,
    'logs': []
}

@app.route('/')
def index():
    """Halaman utama"""
    return render_template('index.html')

@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current bot status"""
    return jsonify(progress_data)

@app.route('/api/start', methods=['POST'])
def start_bot():
    """Start bot process"""
    global bot_instance, bot_thread, is_bot_running, progress_data
    
    data = request.json
    excel_path = data.get('excel_path')
    
    if not excel_path or not os.path.exists(excel_path):
        return jsonify({'error': 'File Excel tidak ditemukan'}), 400
    
    if is_bot_running:
        return jsonify({'error': 'Bot sedang berjalan'}), 400
    
    # Reset progress data
    progress_data = {
        'current': 0,
        'total': 0,
        'success': 0,
        'failed': 0,
        'status': 'Starting bot...',
        'is_success': True,
        'logs': []
    }
    
    # Create bot instance
    bot_instance = LasikBotService()
    bot_instance.excel_path = excel_path
    
    # Set callbacks
    bot_instance.set_on_progress(lambda c, t: update_progress(c, t))
    bot_instance.set_on_status(lambda msg, success: update_status(msg, success))
    
    # Run bot in separate thread
    is_bot_running = True
    bot_thread = threading.Thread(target=run_bot)
    bot_thread.start()
    
    return jsonify({'message': 'Bot started successfully'})

@app.route('/api/stop', methods=['POST'])
def stop_bot():
    """Stop bot process"""
    global bot_instance, is_bot_running
    
    if bot_instance:
        bot_instance.stop()
    
    is_bot_running = False
    update_status("Bot dihentikan oleh user", False)
    
    return jsonify({'message': 'Bot stopped successfully'})

@app.route('/api/download', methods=['GET'])
def download_result():
    """Download result Excel file"""
    if bot_instance and bot_instance.current_output_path and os.path.exists(bot_instance.current_output_path):
        return send_file(
            bot_instance.current_output_path,
            as_attachment=True,
            download_name=os.path.basename(bot_instance.current_output_path)
        )
    return jsonify({'error': 'File tidak ditemukan'}), 404

def run_bot():
    """Run bot in thread"""
    global bot_instance, is_bot_running
    
    try:
        # Open Chrome
        update_status("Membuka Chrome...", True)
        bot_instance.open_chrome_only()
        
        # Attach bot
        update_status("Menghubungkan ke Chrome...", True)
        bot_instance.attach_bot()
        
        # Start processing
        update_status("Memulai proses...", True)
        bot_instance.start()
        
    except Exception as e:
        update_status(f"Error: {str(e)}", False)
    finally:
        is_bot_running = False
        update_status("Bot selesai", True)

def update_progress(current, total):
    """Update progress"""
    global progress_data
    progress_data['current'] = current
    progress_data['total'] = total
    progress_data['success'] = bot_instance.success_count if bot_instance else 0
    progress_data['failed'] = bot_instance.failed_count if bot_instance else 0

def update_status(message, is_success):
    """Update status"""
    global progress_data
    progress_data['status'] = message
    progress_data['is_success'] = is_success
    
    # Add to logs
    log_entry = {
        'timestamp': datetime.now().strftime('%H:%M:%S'),
        'message': message,
        'is_success': is_success
    }
    progress_data['logs'].insert(0, log_entry)  # Add to beginning for latest first
    
    # Keep only last 100 logs
    if len(progress_data['logs']) > 100:
        progress_data['logs'] = progress_data['logs'][:100]

if __name__ == '__main__':
    from datetime import datetime
    app.run(debug=True, port=5000)
