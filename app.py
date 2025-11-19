import os
import time
import logging
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import pandas as pd
import threading
import uuid
from datetime import datetime
from pyxlsb import open_workbook
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import shutil

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

# Garantir que as pastas existam
for folder in [UPLOAD_FOLDER, 'logs', 'templates']:
    os.makedirs(folder, exist_ok=True)

# Dicionário para armazenar o progresso das conversões
conversion_progress = {}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def detect_formatting(value):
    """Detecta formatação baseada no valor da célula"""
    formatting = {
        'font': None,
        'fill': None,
        'alignment': None,
        'border': None,
        'number_format': 'General'
    }
    
    if value is None:
        return formatting
    
    try:
        # Detectar números
        if isinstance(value, (int, float)):
            if isinstance(value, int):
                formatting['number_format'] = '#,##0'
            else:
                # Verificar se é decimal
                if value == int(value):
                    formatting['number_format'] = '#,##0'
                else:
                    formatting['number_format'] = '#,##0.00'
        
        # Detectar texto longo (possível cabeçalho)
        elif isinstance(value, str):
            if len(value) > 20:
                formatting['alignment'] = Alignment(wrap_text=True)
            if value.isupper() or any(word in value.lower() for word in ['total', 'soma', 'quantidade', 'valor']):
                formatting['font'] = Font(bold=True)
                formatting['fill'] = PatternFill(start_color="DDDDDD", fill_type="solid")
    
    except Exception:
        pass
    
    return formatting

def apply_formatting(cell, formatting):
    """Aplica formatação a uma célula"""
    try:
        if formatting.get('font'):
            cell.font = formatting['font']
        if formatting.get('fill'):
            cell.fill = formatting['fill']
        if formatting.get('alignment'):
            cell.alignment = formatting['alignment']
        if formatting.get('number_format'):
            cell.number_format = formatting['number_format']
    except Exception as e:
        logging.debug(f"Erro ao aplicar formatação: {e}")

def convert_xlsb_to_xlsx_advanced(filepath_in, filepath_out, task_id):
    """Conversão avançada que preserva dados e estrutura"""
    try:
        logging.info(f"Iniciando conversão avançada: {filepath_in} -> {filepath_out}")
        
        conversion_progress[task_id] = {
            'status': 'iniciando',
            'progress': 0,
            'message': 'Iniciando conversão...',
            'filename': None,
            'error': None,
            'start_time': datetime.now().isoformat()
        }
        
        # Verificar se arquivo existe
        if not os.path.exists(filepath_in):
            raise FileNotFoundError(f"Arquivo não encontrado: {filepath_in}")
        
        file_size = os.path.getsize(filepath_in)
        conversion_progress[task_id].update({
            'progress': 10,
            'message': f'Arquivo carregado ({file_size / 1024 / 1024:.1f} MB)'
        })
        
        # Método 1: Tentar com pandas + openpyxl
        try:
            conversion_progress[task_id].update({
                'progress': 20,
                'message': 'Lendo estrutura do arquivo XLSB...'
            })
            
            # Ler metadados do arquivo
            xlsb_file = pd.ExcelFile(filepath_in, engine='pyxlsb')
            sheet_names = xlsb_file.sheet_names
            
            conversion_progress[task_id].update({
                'progress': 30,
                'message': f'Encontradas {len(sheet_names)} planilhas'
            })
            
            # Criar workbook de saída
            wb_out = Workbook()
            # Remover sheet padrão
            wb_out.remove(wb_out.active)
            
            # Processar cada planilha
            for sheet_idx, sheet_name in enumerate(sheet_names):
                progress = 30 + (sheet_idx * 60 / len(sheet_names))
                conversion_progress[task_id].update({
                    'progress': progress,
                    'message': f'Processando: {sheet_name}'
                })
                
                try:
                    # Ler dados mantendo tipos originais
                    df = pd.read_excel(
                        filepath_in, 
                        sheet_name=sheet_name, 
                        engine='pyxlsb',
                        dtype=object,
                        keep_default_na=False
                    )
                    
                    # Criar nova planilha
                    ws_out = wb_out.create_sheet(title=sheet_name[:31])
                    
                    # Escrever dados
                    for row_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                        for col_idx, value in enumerate(row, 1):
                            cell = ws_out.cell(row=row_idx, column=col_idx, value=value)
                            
                            # Aplicar formatação detectada
                            formatting = detect_formatting(value)
                            apply_formatting(cell, formatting)
                    
                    # Ajustar largura das colunas
                    for column in ws_out.columns:
                        max_length = 0
                        column_letter = get_column_letter(column[0].column)
                        
                        for cell in column:
                            try:
                                if cell.value:
                                    length = len(str(cell.value))
                                    max_length = max(max_length, length)
                            except:
                                pass
                        
                        adjusted_width = min(max(max_length + 2, 8), 50)
                        ws_out.column_dimensions[column_letter].width = adjusted_width
                    
                    # Adicionar bordas básicas
                    thin_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'), 
                        bottom=Side(style='thin')
                    )
                    
                    for row in ws_out.iter_rows(min_row=1, max_row=ws_out.max_row, min_col=1, max_col=ws_out.max_column):
                        for cell in row:
                            if cell.value is not None:
                                cell.border = thin_border
                    
                    logging.info(f"Planilha {sheet_name} processada com sucesso")
                    
                except Exception as e:
                    logging.error(f"Erro na planilha {sheet_name}: {e}")
                    # Criar planilha vazia como fallback
                    ws_out = wb_out.create_sheet(title=sheet_name[:31])
                    ws_out.cell(1, 1, value=f"Erro ao processar: {str(e)}")
                    continue
                
                time.sleep(0.1)
            
            # Salvar arquivo
            conversion_progress[task_id].update({
                'progress': 95,
                'message': 'Salvando arquivo XLSX...'
            })
            
            wb_out.save(filepath_out)
            
            # Verificar se arquivo foi criado
            if os.path.exists(filepath_out):
                output_size = os.path.getsize(filepath_out)
                conversion_progress[task_id].update({
                    'progress': 100,
                    'message': f'Conversão concluída! ({output_size / 1024 / 1024:.1f} MB)',
                    'status': 'completo',
                    'filename': os.path.basename(filepath_out),
                    'end_time': datetime.now().isoformat()
                })
                logging.info(f"Conversão bem-sucedida: {filepath_out}")
            else:
                raise Exception("Arquivo de saída não foi criado")
                
        except Exception as e:
            logging.error(f"Erro no método principal: {e}")
            
            # Método 2: Fallback simples
            conversion_progress[task_id].update({
                'progress': 50,
                'message': 'Usando método alternativo...'
            })
            
            try:
                xlsb_file = pd.ExcelFile(filepath_in, engine='pyxlsb')
                sheet_names = xlsb_file.sheet_names
                
                with pd.ExcelWriter(filepath_out, engine='openpyxl') as writer:
                    for i, sheet_name in enumerate(sheet_names):
                        df = pd.read_excel(filepath_in, sheet_name=sheet_name, engine='pyxlsb')
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                conversion_progress[task_id].update({
                    'progress': 100,
                    'message': 'Conversão concluída (método simples)',
                    'status': 'completo', 
                    'filename': os.path.basename(filepath_out),
                    'end_time': datetime.now().isoformat()
                })
                
            except Exception as fallback_error:
                raise Exception(f"Todos os métodos falharam: {str(e)} -> {str(fallback_error)}")
        
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
            task_id = str(uuid.uuid4())
            filename = secure_filename(file.filename)
            filepath_in = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath_in)
            
            filename_out = filename.replace('.xlsb', '.xlsx')
            filepath_out = os.path.join(app.config['UPLOAD_FOLDER'], filename_out)
            
            # Informações iniciais
            input_size = os.path.getsize(filepath_in)
            conversion_progress[task_id] = {
                'status': 'iniciando',
                'progress': 0,
                'message': 'Preparando conversão...',
                'filename': filename_out,
                'error': None,
                'start_time': datetime.now().isoformat(),
                'details': {
                    'input_file': filename,
                    'input_size': f"{input_size / 1024 / 1024:.2f} MB"
                }
            }
            
            # Iniciar conversão
            thread = threading.Thread(
                target=convert_xlsb_to_xlsx_advanced,
                args=(filepath_in, filepath_out, task_id)
            )
            thread.daemon = True
            thread.start()
            
            return jsonify({
                'task_id': task_id, 
                'filename': filename_out
            })
        
        return jsonify({'error': 'Tipo de arquivo não permitido'}), 400
    
    except Exception as e:
        logging.error(f"Erro no upload: {e}")
        return jsonify({'error': f'Erro interno: {str(e)}'}), 500

@app.route('/progress/<task_id>')
def get_progress(task_id):
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

# Rota para limpar arquivos antigos
@app.route('/cleanup', methods=['POST'])
def cleanup_files():
    try:
        cutoff_time = time.time() - 3600  # 1 hora
        removed = 0
        
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            if os.path.getctime(filepath) < cutoff_time:
                os.remove(filepath)
                removed += 1
        
        return jsonify({'message': f'{removed} arquivos removidos'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    logging.info("Iniciando aplicação Flask na porta 9090")
    app.run(host='0.0.0.0', port=9090, debug=False)