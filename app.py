import os
import time
import logging
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import pandas as pd
import threading
import uuid
from datetime import datetime

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log'),
        logging.StreamHandler()
    ]
)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsb'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'uma_chave_secreta_muito_segura'

# Garantir que as pastas existem
for folder in [UPLOAD_FOLDER, 'logs', 'static', 'templates']:
    os.makedirs(folder, exist_ok=True)

# Dicionário para armazenar o progresso das conversões
conversion_progress = {}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_xlsb_to_xlsx(filepath_in, filepath_out, task_id):
    """Função de conversão com rastreamento de progresso"""
    try:
        logging.info(f"Iniciando conversão: {filepath_in} -> {filepath_out}")
        
        conversion_progress[task_id] = {
            'status': 'iniciando',
            'progress': 0,
            'message': 'Iniciando conversão...',
            'filename': None,
            'error': None,
            'start_time': datetime.now().isoformat()
        }
        
        # Verificar se arquivo de entrada existe
        if not os.path.exists(filepath_in):
            raise FileNotFoundError(f"Arquivo de entrada não encontrado: {filepath_in}")
        
        # Passo 1: Lendo arquivo XLSB
        conversion_progress[task_id].update({
            'progress': 10,
            'message': 'Lendo arquivo XLSB...'
        })
        
        # Para múltiplas planilhas
        xlsb_file = pd.ExcelFile(filepath_in, engine='pyxlsb')
        sheet_names = xlsb_file.sheet_names
        
        conversion_progress[task_id].update({
            'progress': 20,
            'message': f'Encontradas {len(sheet_names)} planilha(s)'
        })
        
        # Criar writer para XLSX
        with pd.ExcelWriter(filepath_out, engine='openpyxl') as writer:
            for i, sheet_name in enumerate(sheet_names):
                progress = 20 + (i * 70 / len(sheet_names))
                conversion_progress[task_id].update({
                    'progress': progress,
                    'message': f'Convertendo planilha: {sheet_name} ({i+1}/{len(sheet_names)})'
                })
                
                # Ler cada planilha
                df = pd.read_excel(filepath_in, sheet_name=sheet_name, engine='pyxlsb')
                
                # Escrever no arquivo XLSX
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                time.sleep(0.3)
        
        conversion_progress[task_id].update({
            'progress': 95,
            'message': 'Finalizando conversão...',
            'status': 'concluindo'
        })
        
        time.sleep(0.5)
        
        conversion_progress[task_id].update({
            'progress': 100,
            'message': 'Conversão concluída com sucesso!',
            'status': 'completo',
            'filename': os.path.basename(filepath_out),
            'end_time': datetime.now().isoformat()
        })
        
        logging.info(f"Conversão concluída: {filepath_out}")
        
    except Exception as e:
        error_msg = f"Erro na conversão: {str(e)}"
        logging.error(error_msg)
        conversion_progress[task_id].update({
            'status': 'erro',
            'message': error_msg,
            'error': str(e),
            'end_time': datetime.now().isoformat()
        })

@app.route('/')
def index():
    logging.info("Acesso à página principal")
    return render_template('upload.html')

@app.route('/health')
def health_check():
    return jsonify({'status': 'healthy', 'timestamp': datetime.now().isoformat()})

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if file and allowed_file(file.filename):
            # Gerar ID único para a tarefa
            task_id = str(uuid.uuid4())
            
            filename = secure_filename(file.filename)
            filepath_in = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath_in)
            
            filename_out = filename.replace('.xlsb', '.xlsx')
            filepath_out = os.path.join(app.config['UPLOAD_FOLDER'], filename_out)
            
            # Iniciar conversão em thread separada
            thread = threading.Thread(
                target=convert_xlsb_to_xlsx,
                args=(filepath_in, filepath_out, task_id)
            )
            thread.daemon = True
            thread.start()
            
            return jsonify({'task_id': task_id, 'filename': filename_out})
        
        return jsonify({'error': 'Tipo de arquivo não permitido'}), 400
    
    except Exception as e:
        logging.error(f"Erro no upload: {e}")
        return jsonify({'error': f'Erro interno: {str(e)}'}), 500

@app.route('/progress/<task_id>')
def get_progress(task_id):
    """Endpoint para verificar progresso"""
    progress_data = conversion_progress.get(task_id, {
        'status': 'nao_encontrado',
        'progress': 0,
        'message': 'Tarefa não encontrada'
    })
    
    return jsonify(progress_data)

@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_from_directory(
            app.config['UPLOAD_FOLDER'], 
            filename, 
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        logging.error(f"Erro no download: {e}")
        return jsonify({'error': 'Arquivo não encontrado'}), 404

if __name__ == '__main__':
    logging.info("Iniciando aplicação Flask na porta 9090")
    app.run(host='0.0.0.0', port=9090, debug=False)