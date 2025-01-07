import pandas as pd
from xlwt import Workbook, easyxf
from datetime import datetime

# Configuração para estilos de texto (apenas cores de fonte)
style_error = easyxf('font: colour red;')
style_warning = easyxf('font: colour blue;')
style_default = easyxf('font: colour black;')

# Função para processar script-logs
def process_script_logs(filepath):
    data = []
    with open(filepath, 'r') as file:
        for line in file:
            parts = line.split()
            date, time = parts[0], parts[1]
            script_type = parts[2]
            message = ' '.join(parts[3:])
            style = style_error if 'error' in script_type else style_warning
            data.append((date, time, message, script_type, style))
    return data

# Função para processar all-logs
def process_all_logs(filepath):
    data = []
    with open(filepath, 'r') as file:
        for line in file:
            if "netwatch" in line.lower():
                parts = line.split()
                date, time = parts[0], parts[1]
                status = "UP" if "up" in line.lower() else "DOWN"
                operator = "TIM" if "tim" in line.lower() else "CLARO"
                style = style_error if status == "DOWN" else style_warning
                data.append((date, time, operator, status, style))
    return data

# Processar arquivos
script_logs_data = process_script_logs('script-logs.txt.0.txt')
all_logs_data = process_all_logs('all-logs.txt.1.txt')

# Criar workbook para salvar os dados
wb = Workbook()

# Adicionar aba para script-logs
sheet1 = wb.add_sheet('Script Logs')
headers1 = ["Data", "Hora", "Mensagem", "Tipo"]
for col, header in enumerate(headers1):
    sheet1.write(0, col, header, style_default)

for row, (date, time, message, script_type, style) in enumerate(script_logs_data, start=1):
    sheet1.write(row, 0, date, style)
    sheet1.write(row, 1, time, style)
    sheet1.write(row, 2, message, style)
    sheet1.write(row, 3, script_type, style)

# Ajustar largura das colunas na aba Script Logs
for col in range(len(headers1)):
    sheet1.col(col).width = 5000

# Adicionar aba para all-logs
sheet2 = wb.add_sheet('All Logs')
headers2 = ["Data", "Hora", "Operadora", "Status"]
for col, header in enumerate(headers2):
    sheet2.write(0, col, header, style_default)

for row, (date, time, operator, status, style) in enumerate(all_logs_data, start=1):
    sheet2.write(row, 0, date, style)
    sheet2.write(row, 1, time, style)
    sheet2.write(row, 2, operator, style)
    sheet2.write(row, 3, status, style)

# Ajustar largura das colunas na aba All Logs
for col in range(len(headers2)):
    sheet2.col(col).width = 5000

# Adicionar data e hora de geração no arquivo
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
sheet1.write(len(script_logs_data) + 2, 0, f"Gerado em: {timestamp}", style_default)
sheet2.write(len(all_logs_data) + 2, 0, f"Gerado em: {timestamp}", style_default)

# Salvar o arquivo Excel com data/hora no nome
output_file = f"V3_saida_logs_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xls"
wb.save(output_file)
print(f"Arquivo Excel gerado: {output_file}")
