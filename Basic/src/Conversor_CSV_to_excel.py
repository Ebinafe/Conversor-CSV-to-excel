"""
Conversor de CSV para Excel
Descrição: Script simples para converter arquivos CSV em Excel (.xlsx)
"""

import pandas as pd
import os


def converter_csv_para_excel(arquivo_csv, arquivo_excel=None):
    """
    Converte um arquivo CSV para Excel.
    
    Args:
        arquivo_csv (str): Caminho do arquivo CSV de entrada
        arquivo_excel (str, optional): Caminho do arquivo Excel de saída.
                                      Se não informado, usa o mesmo nome do CSV.
    
    Returns:
        bool: True se conversão foi bem-sucedida, False caso contrário
    """
    try:
        # Verifica se o arquivo CSV existe
        if not os.path.exists(arquivo_csv):
            print(f"❌ Erro: Arquivo '{arquivo_csv}' não encontrado!")
            return False
        
        # Define o nome do arquivo Excel de saída
        if arquivo_excel is None:
            arquivo_excel = arquivo_csv.replace('.csv', '.xlsx')
        
        # Lê o arquivo CSV
        print(f"📄 Lendo arquivo CSV: {arquivo_csv}")
        df = pd.read_csv(arquivo_csv, encoding='utf-8')
        
        # Salva como Excel
        print(f"💾 Salvando como Excel: {arquivo_excel}")
        df.to_excel(arquivo_excel, index=False, engine='openpyxl')
        
        print(f"✅ Conversão concluída com sucesso!")
        print(f"📊 Total de linhas: {len(df)}")
        print(f"📊 Total de colunas: {len(df.columns)}")
        
        return True
        
    except UnicodeDecodeError:
        # Tenta com outra codificação se UTF-8 falhar
        try:
            print("⚠️  Tentando com codificação 'latin-1'...")
            df = pd.read_csv(arquivo_csv, encoding='latin-1')
            df.to_excel(arquivo_excel, index=False, engine='openpyxl')
            print(f"✅ Conversão concluída com sucesso!")
            return True
        except Exception as e:
            print(f"❌ Erro ao converter: {e}")
            return False
    
    except Exception as e:
        print(f"❌ Erro ao converter: {e}")
        return False


def main():
    """Função principal do programa"""
    print("=" * 50)
    print("CONVERSOR DE CSV PARA EXCEL")
    print("=" * 50)
    
    # Solicita o nome do arquivo CSV
    arquivo_csv = input("\n📂 Digite o nome do arquivo CSV: ").strip()
    
    # Adiciona extensão .csv se não foi informada
    if not arquivo_csv.endswith('.csv'):
        arquivo_csv += '.csv'
    
    # Pergunta se quer personalizar o nome do arquivo de saída
    usar_nome_personalizado = input("\n❓ Deseja personalizar o nome do arquivo Excel? (s/n): ").strip().lower()
    
    arquivo_excel = None
    if usar_nome_personalizado == 's':
        arquivo_excel = input("📂 Digite o nome do arquivo Excel de saída: ").strip()
        if not arquivo_excel.endswith('.xlsx'):
            arquivo_excel += '.xlsx'
    
    # Realiza a conversão
    print("\n🔄 Iniciando conversão...\n")
    converter_csv_para_excel(arquivo_csv, arquivo_excel)
    
    print("\n" + "=" * 50)


if __name__ == "__main__":
    main()