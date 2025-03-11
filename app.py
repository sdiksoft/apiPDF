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

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'  # Pasta para PDFs gerados
app.config['TEMP_FOLDER'] = 'temp'  # Nova pasta para arquivos XLSX temporários
app.config['MODEL_INFO'] = {}  # Armazena informações dos modelos

def cleanup_temp_files():
    """Limpa arquivos temporários antigos"""
    try:
        if os.path.exists(app.config['TEMP_FOLDER']):
            shutil.rmtree(app.config['TEMP_FOLDER'])
        os.makedirs(app.config['TEMP_FOLDER'])
    except Exception as e:
        print(f"Erro ao limpar arquivos temporários: {str(e)}")

def convert_to_pdf(excel_path, pdf_path):
    """Converte arquivo Excel para PDF usando LibreOffice/soffice"""
    try:
        # Determina o comando do LibreOffice baseado no sistema operacional
        if platform.system() == 'Windows':
            # No Windows, procura o LibreOffice em locais comuns
            program_files = os.environ.get('PROGRAMFILES', 'C:\\Program Files')
            program_files_x86 = os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)')
            possible_paths = [
                os.path.join(program_files, 'LibreOffice', 'program', 'soffice.exe'),
                os.path.join(program_files_x86, 'LibreOffice', 'program', 'soffice.exe'),
                os.path.join(program_files, 'LibreOffice*', 'program', 'soffice.exe'),
                'soffice.exe'  # Se estiver no PATH
            ]
            
            soffice_path = None
            for path in possible_paths:
                if '*' in path:
                    # Procura por pastas que correspondam ao padrão
                    import glob
                    matches = glob.glob(path)
                    if matches:
                        soffice_path = matches[0]
                        break
                elif os.path.exists(path):
                    soffice_path = path
                    break
            
            if not soffice_path:
                raise Exception("LibreOffice não encontrado. Por favor, instale-o ou adicione ao PATH")
            
            command = [soffice_path]
        else:
            # No Linux, assume que 'soffice' está no PATH
            command = ['soffice']
        
        # Adiciona os argumentos comuns
        command.extend([
            '--headless',
            '--convert-to',
            'pdf',
            '--outdir',
            os.path.dirname(pdf_path),
            excel_path
        ])
        
        # Executa o comando
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        
        # Aguarda a conclusão do processo
        stdout, stderr = process.communicate()
        
        if process.returncode != 0:
            raise Exception(f"Erro na conversão: {stderr.decode()}")
        
        # Renomeia o arquivo gerado para o nome desejado
        generated_pdf = Path(excel_path).with_suffix('.pdf')
        if generated_pdf.exists():
            os.rename(generated_pdf, pdf_path)
            return True
        
        raise Exception("Arquivo PDF não foi gerado")
        
    except Exception as e:
        raise Exception(f"Erro ao converter para PDF: {str(e)}")

def analyze_excel_file(filename):
    """Analisa um arquivo Excel para encontrar variáveis e tabelas"""
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    wb = load_workbook(filepath)
    sheet = wb.active
    
    variables = []
    tables = []
    
    # Procura por variáveis e tabelas em todas as células
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                # Procura por variáveis (var.nome.tipo)
                var_matches = re.finditer(r'var\.(\w+)\.(text|date|int|double)', cell.value)
                for match in var_matches:
                    variables.append({
                        'name': match.group(1),
                        'type': match.group(2),
                        'cell': f'{cell.column_letter}{cell.row}'
                    })
                
                # Procura por tabelas (tabela.nome.campo.tipo)
                table_matches = re.finditer(r'tabela\.(\w+)\.(\w+)\.(text|date|int|double)', cell.value)
                for match in table_matches:
                    tables.append({
                        'name': match.group(1),
                        'field': match.group(2),
                        'type': match.group(3),
                        'start_cell': f'{cell.column_letter}{cell.row}'
                    })
    
    return {
        'variables': variables,
        'tables': tables
    }

def create_endpoint_schema(model_info):
    """Cria um exemplo prático de payload baseado nas variáveis e tabelas encontradas"""
    example = {}
    
    # Adiciona variáveis simples com valores de exemplo
    for var in model_info['variables']:
        if var['type'] == 'text':
            example[var['name']] = f"Exemplo {var['name']}"
        elif var['type'] == 'date':
            example[var['name']] = datetime.now().strftime('%d-%m-%Y')
        elif var['type'] == 'int':
            example[var['name']] = 1
        elif var['type'] == 'double':
            example[var['name']] = 1.5
    
    # Adiciona tabelas com exemplos
    table_names = set(table['name'] for table in model_info['tables'])
    for table_name in table_names:
        table_example = {}
        for table in model_info['tables']:
            if table['name'] == table_name:
                if table['type'] == 'text':
                    table_example[table['field']] = f"Exemplo {table['field']}"
                elif table['type'] == 'date':
                    table_example[table['field']] = datetime.now().strftime('%d-%m-%Y')
                elif table['type'] == 'int':
                    table_example[table['field']] = 1
                elif table['type'] == 'double':
                    table_example[table['field']] = 1.5
        
        example[table_name] = [table_example]  # Array com um item de exemplo
    
    return example

def generate_error_pdf(error_message=None):
    """Gera um PDF com mensagem de erro"""
    try:
        # Lê o template HTML
        with open('templates/error_template.html', 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        # Renderiza o template com a data/hora atual
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        html_content = template.render(
            timestamp=timestamp,
            error_message=error_message
        )
        
        # Gera um nome único para o arquivo
        error_filename = f'error_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        error_path = os.path.join(app.config['DOWNLOAD_FOLDER'], error_filename)
        
        # Gera o PDF
        HTML(string=html_content).write_pdf(error_path)
        
        return error_filename
    except Exception as e:
        print(f"Erro ao gerar PDF de erro: {str(e)}")
        return None

# Criar pastas necessárias se não existirem
for folder in [app.config['UPLOAD_FOLDER'], app.config['DOWNLOAD_FOLDER'], app.config['TEMP_FOLDER']]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Limpa arquivos temporários ao iniciar a aplicação
cleanup_temp_files()

@app.route('/')
def index():
    # Listar todos os arquivos xlsx e seus endpoints
    files = []
    for f in os.listdir(app.config['UPLOAD_FOLDER']):
        if f.endswith('.xlsx'):
            model_name = os.path.splitext(f)[0]
            # Constrói a URL completa do endpoint
            endpoint = request.host_url.rstrip('/') + f'/api/generate/{model_name}'
            if f not in app.config['MODEL_INFO']:
                app.config['MODEL_INFO'][f] = analyze_excel_file(f)
            
            model_info = app.config['MODEL_INFO'][f]
            example = create_endpoint_schema(model_info)
            
            files.append({
                'name': f,
                'endpoint': endpoint,
                'schema': example
            })
    return render_template('index.html', files=files)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'Nenhum arquivo selecionado', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'Nenhum arquivo selecionado', 400
    
    if file and file.filename.endswith('.xlsx'):
        filename = file.filename
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Analisa o arquivo após o upload
        try:
            app.config['MODEL_INFO'][filename] = analyze_excel_file(filename)
            return 'Arquivo enviado com sucesso!'
        except Exception as e:
            os.remove(filepath)
            return f'Erro ao analisar arquivo: {str(e)}', 400
    
    return 'Tipo de arquivo não permitido', 400

@app.route('/download/<path:filename>')
def download_file(filename):
    # Verifica primeiro na pasta de downloads (PDFs)
    if os.path.exists(os.path.join(app.config['DOWNLOAD_FOLDER'], filename)):
        return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename)
    # Se não encontrar, procura na pasta de uploads (XLSXs)
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

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
        for table_name in set(t['name'] for t in model_info['tables']):
            if table_name in data:
                table_data = data[table_name]
                table_fields = [t for t in model_info['tables'] if t['name'] == table_name]
                
                # Encontra a primeira célula da tabela
                start_cell = table_fields[0]['start_cell']
                row = int(''.join(filter(str.isdigit, start_cell)))
                
                # Insere os dados da tabela
                for i, item in enumerate(table_data):
                    current_row = row + i
                    for field in table_fields:
                        col = ''.join(filter(str.isalpha, field['start_cell']))
                        cell = f'{col}{current_row}'
                        
                        value = item.get(field['field'])
                        if value is not None:
                            if field['type'] == 'date':
                                try:
                                    value = datetime.strptime(value, '%d-%m-%Y')
                                except ValueError:
                                    try:
                                        value = datetime.strptime(value, '%Y-%m-%d')
                                    except ValueError:
                                        return jsonify({'error': f'Formato de data inválido para {field["field"]} em {table_name}. Use DD-MM-YYYY'}), 400
                            elif field['type'] == 'int':
                                value = int(value)
                            elif field['type'] == 'double':
                                value = float(value)
                            
                            sheet[cell] = value
        
        # Gera nomes únicos para os arquivos
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f'generated_{model_name}_{timestamp}.xlsx'
        pdf_filename = f'generated_{model_name}_{timestamp}.pdf'
        
        excel_path = os.path.join(app.config['TEMP_FOLDER'], excel_filename)
        pdf_path = os.path.join(app.config['DOWNLOAD_FOLDER'], pdf_filename)
        
        # Salva o arquivo Excel temporário
        wb.save(excel_path)
        
        try:
            # Tenta converter para PDF
            convert_to_pdf(excel_path, pdf_path)
            
            # Se a conversão for bem-sucedida, remove o arquivo Excel temporário
            if os.path.exists(excel_path):
                os.remove(excel_path)
            
            return jsonify({
                'message': 'Arquivo PDF gerado com sucesso',
                'download_url': f'/download/{pdf_filename}'
            })
            
        except Exception as e:
            # Se falhar a conversão para PDF
            error_pdf = generate_error_pdf(str(e))
            
            # Move o Excel para a pasta de downloads para disponibilizá-lo
            excel_download_path = os.path.join(app.config['DOWNLOAD_FOLDER'], excel_filename)
            shutil.move(excel_path, excel_download_path)
            
            return jsonify({
                'message': 'Arquivo gerado com sucesso (apenas Excel - falha na conversão para PDF)',
                'error': str(e),
                'download_url': f'/download/{excel_filename}',
                'error_pdf': f'/download/{error_pdf}' if error_pdf else None
            })
    
    except Exception as e:
        # Gera PDF de erro para qualquer outra exceção
        error_pdf = generate_error_pdf(str(e))
        return jsonify({
            'error': str(e),
            'error_pdf': f'/download/{error_pdf}' if error_pdf else None
        }), 500

# Adiciona limpeza de arquivos temporários quando a aplicação é encerrada
import atexit
atexit.register(cleanup_temp_files)

if __name__ == '__main__':
    app.run(debug=True) 