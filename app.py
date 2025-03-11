from flask import Flask, render_template, request, send_from_directory, jsonify, url_for
import os
import pandas as pd
import re
from openpyxl import load_workbook
from datetime import datetime
import json
from pathlib import Path
import subprocess
import platform
from weasyprint import HTML
from jinja2 import Template
import shutil
import threading
from queue import Queue
import time

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'  # Pasta para PDFs gerados
app.config['TEMP_FOLDER'] = 'temp'  # Pasta para arquivos XLSX processados
app.config['MODEL_INFO'] = {}  # Armazena informações dos modelos
app.config['CONVERSION_QUEUE'] = Queue()  # Fila para conversão de PDFs
app.config['CONVERSION_STATUS'] = {}  # Status das conversões

# Inicia o worker de conversão em background
def pdf_conversion_worker():
    while True:
        try:
            # Obtém o próximo item da fila
            excel_path, pdf_path = app.config['CONVERSION_QUEUE'].get()
            
            # Atualiza o status
            conversion_id = os.path.basename(pdf_path)
            app.config['CONVERSION_STATUS'][conversion_id] = {
                'status': 'processing',
                'message': 'Convertendo para PDF...'
            }
            
            try:
                # Tenta converter usando LibreOffice
                if platform.system() == 'Linux':
                    # Verifica os possíveis caminhos do LibreOffice no Ubuntu
                    soffice_paths = [
                        '/usr/bin/soffice',
                        '/usr/bin/libreoffice',
                        '/usr/lib/libreoffice/program/soffice',
                    ]
                    
                    soffice = None
                    for path in soffice_paths:
                        if os.path.exists(path):
                            soffice = path
                            break
                    
                    if soffice:
                        # Mata qualquer processo do LibreOffice que possa estar rodando
                        try:
                            subprocess.run(['pkill', 'soffice'], stderr=subprocess.DEVNULL)
                            time.sleep(1)  # Espera um segundo para garantir que o processo foi finalizado
                        except:
                            pass
                        
                        # Converte para PDF
                        process = subprocess.run([
                            soffice,
                            '--headless',
                            '--convert-to', 'pdf',
                            '--outdir', os.path.dirname(pdf_path),
                            excel_path
                        ], check=True, timeout=30)  # 30 segundos de timeout
                        
                        # Renomeia o arquivo
                        temp_pdf = os.path.join(os.path.dirname(pdf_path), 
                                              os.path.splitext(os.path.basename(excel_path))[0] + '.pdf')
                        if os.path.exists(temp_pdf):
                            shutil.move(temp_pdf, pdf_path)
                            app.config['CONVERSION_STATUS'][conversion_id] = {
                                'status': 'completed',
                                'message': 'Conversão concluída com sucesso',
                                'pdf_url': f'/download/{os.path.basename(pdf_path)}'
                            }
                        else:
                            raise Exception("PDF não foi gerado")
                    else:
                        raise Exception("LibreOffice não encontrado. Instale com: sudo apt-get install libreoffice")
                else:
                    raise Exception("Sistema operacional não suportado")
                    
            except subprocess.TimeoutExpired:
                app.config['CONVERSION_STATUS'][conversion_id] = {
                    'status': 'error',
                    'message': 'Tempo limite excedido na conversão do PDF'
                }
            except Exception as e:
                app.config['CONVERSION_STATUS'][conversion_id] = {
                    'status': 'error',
                    'message': f'Erro na conversão: {str(e)}'
                }
            
            # Remove arquivos temporários após 1 hora
            threading.Timer(3600, cleanup_temp_files, args=[excel_path, pdf_path]).start()
            
        except Exception as e:
            print(f"Erro no worker de conversão: {str(e)}")
        finally:
            app.config['CONVERSION_QUEUE'].task_done()

def cleanup_temp_files(excel_path, pdf_path):
    """Remove arquivos temporários após um período"""
    try:
        if os.path.exists(excel_path):
            os.remove(excel_path)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
    except Exception as e:
        print(f"Erro ao limpar arquivos temporários: {str(e)}")

# Inicia o thread do worker
conversion_thread = threading.Thread(target=pdf_conversion_worker, daemon=True)
conversion_thread.start()

@app.route('/conversion-status/<conversion_id>')
def conversion_status(conversion_id):
    """Retorna o status atual da conversão"""
    status = app.config['CONVERSION_STATUS'].get(conversion_id, {
        'status': 'not_found',
        'message': 'Conversão não encontrada'
    })
    return jsonify(status)

@app.route('/')
def index():
    # Lista todos os arquivos XLSX na pasta de uploads
    files = []
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            if filename.endswith('.xlsx'):
                file_info = {
                    'name': filename,
                    'endpoint': url_for('generate_from_model', model_name=filename[:-5], _external=True),
                    'schema': app.config['MODEL_INFO'].get(filename, {})
                }
                files.append(file_info)
    return render_template('index.html', files=files)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'Nenhum arquivo enviado', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'Nenhum arquivo selecionado', 400
    
    if not file.filename.endswith('.xlsx'):
        return 'Apenas arquivos XLSX são permitidos', 400

    try:
        # Garante que a pasta uploads existe
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        
        # Salva o arquivo
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        
        # Analisa o arquivo XLSX para extrair informações
        wb = load_workbook(filepath)
        sheet = wb.active
        
        # Inicializa as informações do modelo
        model_info = {
            'variables': [],
            'tables': []
        }
        
        # Procura por células com marcadores especiais
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    # Procura por variáveis (formato: ${nome:tipo})
                    var_matches = re.finditer(r'\${([^:]+):([^}]+)}', cell.value)
                    for match in var_matches:
                        name, type_info = match.groups()
                        model_info['variables'].append({
                            'name': name,
                            'type': type_info,
                            'cell': cell.coordinate
                        })
                    
                    # Procura por tabelas (formato: #{tabela.campo:tipo})
                    table_matches = re.finditer(r'#{([^.]+)\.([^:]+):([^}]+)}', cell.value)
                    for match in table_matches:
                        table_name, field, type_info = match.groups()
                        model_info['tables'].append({
                            'name': table_name,
                            'field': field,
                            'type': type_info,
                            'start_cell': cell.coordinate
                        })
        
        # Armazena as informações do modelo
        app.config['MODEL_INFO'][file.filename] = model_info
        
        return 'Arquivo enviado com sucesso', 200
    
    except Exception as e:
        return f'Erro ao processar arquivo: {str(e)}', 500

@app.route('/download/<path:filename>')
def download_file(filename):
    # Verifica em todas as pastas possíveis
    for folder in [app.config['DOWNLOAD_FOLDER'], app.config['TEMP_FOLDER'], app.config['UPLOAD_FOLDER']]:
        if os.path.exists(os.path.join(folder, filename)):
            return send_from_directory(folder, filename)
    return "Arquivo não encontrado", 404

@app.route('/delete/<filename>', methods=['POST'])
def delete_model(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            if filename in app.config['MODEL_INFO']:
                del app.config['MODEL_INFO'][filename]
            return jsonify({'message': 'Modelo excluído com sucesso'}), 200
        return jsonify({'error': 'Arquivo não encontrado'}), 404
    except Exception as e:
        return jsonify({'error': f'Erro ao excluir arquivo: {str(e)}'}), 500

@app.route('/api/generate/<model_name>', methods=['POST'])
def generate_from_model(model_name):
    filename = f'{model_name}.xlsx'
    if filename not in app.config['MODEL_INFO']:
        error_pdf = generate_error_pdf("Modelo não encontrado")
        if error_pdf:
            return jsonify({
                'error': 'Modelo não encontrado',
                'error_pdf': f'/download/{error_pdf}'
            }), 404
        return jsonify({'error': 'Modelo não encontrado'}), 404
    
    try:
        # Carrega o arquivo modelo
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        wb = load_workbook(template_path)
        sheet = wb.active
        
        data = request.get_json()
        model_info = app.config['MODEL_INFO'][filename]
        
        # Substitui variáveis simples
        for var in model_info['variables']:
            if var['name'] in data:
                cell = sheet[var['cell']]
                value = data[var['name']]
                
                # Converte o valor para o tipo apropriado
                if var['type'] == 'date':
                    try:
                        # Tenta primeiro o formato DD-MM-YYYY
                        value = datetime.strptime(value, '%d-%m-%Y')
                    except ValueError:
                        try:
                            # Se falhar, tenta o formato YYYY-MM-DD
                            value = datetime.strptime(value, '%Y-%m-%d')
                        except ValueError:
                            return jsonify({'error': f'Formato de data inválido para {var["name"]}. Use DD-MM-YYYY'}), 400
                elif var['type'] == 'int':
                    value = int(value)
                elif var['type'] == 'double':
                    value = float(value)
                
                cell.value = value
        
        # Processa tabelas
        table_positions = {}  # Armazena a última linha usada para cada tabela
        
        # Primeiro, organiza os campos por tabela
        tables = {}
        for table_info in model_info['tables']:
            table_name = table_info['name']
            if table_name not in tables:
                tables[table_name] = []
            tables[table_name].append(table_info)
        
        # Processa cada tabela
        for table_name, table_fields in tables.items():
            if table_name in data:
                table_data = data[table_name]
                if not isinstance(table_data, list):
                    return jsonify({'error': f'Dados da tabela {table_name} devem ser uma lista'}), 400
                
                # Encontra a linha inicial da tabela
                start_row = min(int(''.join(filter(str.isdigit, field['start_cell']))) for field in table_fields)
                current_row = start_row + 1  # Começa na linha após os cabeçalhos
                
                # Para cada item na lista de dados
                for item in table_data:
                    # Para cada campo da tabela
                    for field in table_fields:
                        # Obtém a coluna da célula original
                        col = ''.join(filter(str.isalpha, field['start_cell']))
                        cell = f'{col}{current_row}'
                        
                        # Obtém o valor do campo
                        value = item.get(field['field'])
                        if value is not None:
                            # Converte o valor para o tipo apropriado
                            if field['type'] == 'date':
                                try:
                                    value = datetime.strptime(value, '%d-%m-%Y')
                                except ValueError:
                                    try:
                                        value = datetime.strptime(value, '%Y-%m-%d')
                                    except ValueError:
                                        return jsonify({'error': f'Formato de data inválido para {field["field"]} em {table_name}'}), 400
                            elif field['type'] == 'int':
                                value = int(value)
                            elif field['type'] == 'double':
                                value = float(value)
                            
                            # Atribui o valor à célula
                            sheet[cell] = value
                    
                    current_row += 1  # Move para a próxima linha
        
        # Gera nomes únicos para os arquivos
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f'generated_{model_name}_{timestamp}.xlsx'
        pdf_filename = f'generated_{model_name}_{timestamp}.pdf'
        
        # Garante que as pastas existem
        os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)
        os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
        
        # Salva o arquivo Excel temporário
        excel_path = os.path.join(app.config['TEMP_FOLDER'], excel_filename)
        wb.save(excel_path)
        
        # Inicia a conversão para PDF em background
        pdf_path = os.path.join(app.config['DOWNLOAD_FOLDER'], pdf_filename)
        app.config['CONVERSION_QUEUE'].put((excel_path, pdf_path))
        
        return jsonify({
            'message': 'Arquivo gerado com sucesso',
            'excel_url': f'/download/{excel_filename}',
            'conversion_id': pdf_filename,
            'status_url': f'/conversion-status/{pdf_filename}'
        })
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)