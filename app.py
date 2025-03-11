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
app.config['TEMP_FOLDER'] = 'temp'  # Pasta para arquivos XLSX processados
app.config['MODEL_INFO'] = {}  # Armazena informações dos modelos

# ... rest of the code ...

# Criar cometario de verssão do codigo verssao 1.1  


@app.route('/download/<path:filename>')
def download_file(filename):
    # Verifica em todas as pastas possíveis
    for folder in [app.config['DOWNLOAD_FOLDER'], app.config['TEMP_FOLDER'], app.config['UPLOAD_FOLDER']]:
        if os.path.exists(os.path.join(folder, filename)):
            return send_from_directory(folder, filename)
    return "Arquivo não encontrado", 404

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
        
        # Salva o XLSX na pasta temp
        excel_path = os.path.join(app.config['TEMP_FOLDER'], excel_filename)
        pdf_path = os.path.join(app.config['DOWNLOAD_FOLDER'], pdf_filename)
        
        # Salva o arquivo Excel
        wb.save(excel_path)
        
        try:
            # Tenta converter para PDF em background
            convert_to_pdf(excel_path, pdf_path)
            
            # Retorna ambos os links, já que o PDF pode demorar para ser gerado
            return jsonify({
                'message': 'Arquivo gerado com sucesso',
                'excel_url': f'/download/{excel_filename}',
                'pdf_url': f'/download/{pdf_filename}'
            })
            
        except Exception as e:
            # Se falhar a conversão para PDF
            error_pdf = generate_error_pdf(str(e))
            
            return jsonify({
                'message': 'Arquivo gerado com sucesso (falha na conversão para PDF)',
                'error': str(e),
                'excel_url': f'/download/{excel_filename}',
                'error_pdf': f'/download/{error_pdf}' if error_pdf else None
            })
    
    except Exception as e:
        # Gera PDF de erro para qualquer outra exceção
        error_pdf = generate_error_pdf(str(e))
        return jsonify({
            'error': str(e),
            'error_pdf': f'/download/{error_pdf}' if error_pdf else None
        }), 500

if __name__ == '__main__':
    app.run(debug=True)