import subprocess
import socket
import pandas as pd
from datetime import datetime
import platform
import time
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import PieChart, BarChart, Reference

# Função para verificar DHCP apenas uma vez.
arp_output = None

def get_arp_output():
    global arp_output
    if arp_output is None:
        try:
            result = subprocess.run(['arp', '-a'], stdout=subprocess.PIPE, text=True)
            arp_output = result.stdout
            print(arp_output)  # Exibe o resultado completo no terminal
        except Exception as e:
            print(f"Erro ao verificar DHCP: {e}")
    return arp_output

# Função para executar o ping em um IP e retornar o status, latência e contagem de pacotes.
def ping_ip(ip):
    try:
        param = '-n' if platform.system().lower() == 'windows' else '-c'
        output = subprocess.run(['ping', param, '4', '-w', '10', ip], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        if output.returncode == 0:
            latencies = [int(line.split('tempo=')[1].split('ms')[0]) for line in output.stdout.splitlines() if 'tempo=' in line]
            if latencies:
                min_latency = min(latencies)
                max_latency = max(latencies)
                avg_latency = sum(latencies) // len(latencies)
                print(output.stdout)  # Exibe o resultado do ping no terminal
                return 'Ativo', avg_latency, min_latency, max_latency, 4, 4, output.stdout  
            else:
                return 'Ativo', 'Latência indisponível', 0, 0, 4, 4, output.stdout
        else:
            return 'Inativo', None, 0, 0, 0, 0, output.stdout
    except Exception as e:
        return 'Inativo', None, 0, 0, 0, 0, str(e)

# Função para obter o hostname a partir de um IP scaneado na rede (obs: Função ainda em teste pois ainda gera alguns erros.).
def get_hostname(ip):
    try:
        hostname = socket.gethostbyaddr(ip)[0]
        return hostname
    except (socket.herror, socket.gaierror):
        return 'x'

# Função para escanear um único IP e retornar suas informações, desenvolvido a partir de comandos do CMD .
def scan_ip(ip):
    status, avg_latency, min_latency, max_latency, packets_sent, packets_received, ping_output = ping_ip(ip)
    hostname = get_hostname(ip) if status == 'Ativo' else 'x'
    arp_output = get_arp_output()  # Coleta informações da rede e retorna os valores esptipulados para mim 
    return {
        'IP': ip,
        'Hostname': hostname,
        'Status': status,
        'Latência Média (ms)': avg_latency if avg_latency is not None else 'N/A',
        'Latência Mínima (ms)': min_latency,
        'Latência Máxima (ms)': max_latency,
        'Pacotes Enviados': packets_sent,
        'Pacotes Recebidos': packets_received,
        'ARP Output': arp_output
    }

# Função para escanear a rede e coletar informações dos IPs.
def scan_network(ip_range):
    data = []
    for ip in ip_range:
        print(f"Escaneando IP: {ip}...")  # Mensagem de escaneamento, somente para vizualização de que está sendo feito.
        status, avg_latency, min_latency, max_latency, packets_sent, packets_received, ping_output = ping_ip(ip)
        if status == 'Ativo':
            time.sleep(4)  #Estipulei um tempo para que seja realizado o proximo ping 
            arp_output = get_arp_output()
            time.sleep(4)  # Espera 10 segundos apos verificar as informações que busca na rede
            hostname = get_hostname(ip) if status == 'Ativo' else 'x'
            data.append({
                'IP': ip,
                'Hostname': hostname,
                'Status': status,
                'Latência Média (ms)': avg_latency if avg_latency is not None else 'N/A',
                'Latência Mínima (ms)': min_latency,
                'Latência Máxima (ms)': max_latency,
                'Pacotes Enviados': packets_sent,
                'Pacotes Recebidos': packets_received,
                'ARP Output': arp_output
            })
            print(f"Resultado para o IP: {ip} - Status: {status} - Hostname: {hostname} - Latência Média: {avg_latency} ms")
        else:
            data.append({
                'IP': ip,
                'Hostname': 'x',
                'Status': 'Inativo',
                'Latência Média (ms)': 'N/A',
                'Latência Mínima (ms)': 'N/A',
                'Latência Máxima (ms)': 'N/A',
                'Pacotes Enviados': 0,
                'Pacotes Recebidos': 0,
                'ARP Output': 'N/A'
            })
    return data

# Função para analisar os dados coletados e gerar estatísticas sobre máquinas ativas e inativas.
def analyze_data(data):
    active_count = sum(1 for entry in data if entry['Status'] == 'Ativo')
    inactive_count = sum(1 for entry in data if entry['Status'] == 'Inativo')
    
    print(f"\nAnálise de dados:")
    print(f"Total de máquinas ativas: {active_count}")
    print(f"Total de máquinas inativas: {inactive_count}")
    
    return active_count, inactive_count

# Função para coletar informações adicionais da rede usando netstat e route.
def get_network_info():
    try:
        netstat_output = subprocess.run(['netstat', '-a'], stdout=subprocess.PIPE, text=True)
        route_output = subprocess.run(['route', 'print'], stdout=subprocess.PIPE, text=True)
        return netstat_output.stdout, route_output.stdout
    except Exception as e:
        print(f"Erro ao coletar informações da rede: {e}")
        return 'N/A', 'N/A'

# Função para salvar os dados coletados em um arquivo Excel e criar gráficos.
def save_to_excel(data, active_count, inactive_count):
    df = pd.DataFrame(data)
    directory = r'C:\Users\administrator\Desktop\analise de dados de redes'
    filename = f'{directory}\rede_status_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    try:
        df.to_excel(filename, index=False)
        print(f"Arquivo salvo: {filename}")
        
        wb = openpyxl.load_workbook(filename)
        ws = wb.active

        pie_chart = PieChart()
        pie_chart_data = Reference(ws, min_col=3, min_row=1, max_row=2)
        pie_chart.add_data(pie_chart_data, titles_from_data=True)
        pie_chart.title = "Distribuição de Máquinas Ativas e Inativas"
        ws.add_chart(pie_chart, "E5")  

        bar_chart = BarChart()
        bar_chart_data = Reference(ws, min_col=2, min_row=1, max_row=len(data) + 1, max_col=8)
        bar_chart.add_data(bar_chart_data, titles_from_data=True)
        bar_chart.title = "Status de Rede"
        ws.add_chart(bar_chart, "E20")  

        wb.save(filename)
        print(f"Gráficos adicionados e arquivo salvo: {filename}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")

# Função para criar intervalos de IPs para escaneamento.
def create_ip_ranges():
    ip_ranges = []
    for i in range(1, 255):
        ip_ranges.append(f'10.85.193.{i}')
    for i in range(1, 200):
        ip_ranges.append(f'172.16.50.{i}')
    for i in range.append(f'172.16.53.{i}')
        ip_ranges.appende(1,10)
    for i in range.append(f'172.16.52.{i}')
        ip_ranges.append(1)
    return ip_ranges

# Função principal que executa o escaneamento de rede em loop.
def main():
    while True:
        ip_ranges = create_ip_ranges()

        print("Iniciando escaneamento da rede 10.85.193.x...")
        data_10 = scan_network(ip_ranges[:254]) 

        print("Iniciando escaneamento da rede 172.16.50.x...")
        data_172 = scan_network(ip_ranges[254:])  

        data = data_10 + data_172
        active_count, inactive_count = analyze_data(data)

        netstat_output, route_output = get_network_info()
        data.append({'IP': 'Informações Adicionais', 'Hostname': 'N/A', 'Status': 'N/A', 
                      'Latência Média (ms)': netstat_output, 'Latência Mínima (ms)': route_output,
                      'Latência Máxima (ms)': 'N/A', 'Pacotes Enviados': 'N/A', 
                      'Pacotes Recebidos': 'N/A', 'ARP Output': 'N/A'})

        save_to_excel(data, active_count, inactive_count)

        print("Aguardando 60 minutos para o próximo escaneamento...")
        time.sleep(3600)  # Espera 60 minutos antes de escanear novamente

# Inicia o script
if __name__ == '__main__':
    main()
