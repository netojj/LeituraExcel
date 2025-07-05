import pandas as pd
import re
import os


def ler_dados_excel(caminho_arquivo: str, nome_aba: str) -> list[dict]:
    print(f"\n--- Iniciando leitura da planilha com Pandas (MODO TESTE: 5 PRIMEIROS) ---")
    print(f"Arquivo: {caminho_arquivo}")
    print(f"Aba: {nome_aba}")
    print("------------------------------------------------------------------------\n")

    dados_planilha = []
    try:
        # Determina o motor de leitura com base na extensão do arquivo
        engine = 'pyxlsb' if caminho_arquivo.endswith('.xlsb') else 'openpyxl'
        df = pd.read_excel(
            caminho_arquivo,
            sheet_name=nome_aba,
            engine=engine,
            header=None,
            keep_default_na=False,
        )

        # Itera sobre as linhas do DataFrame, começando da segunda linha (índice 1)
        for i, row in df.iloc[1:].iterrows():
            protocolo = row.get(0, '')  # Coluna A (índice 0)
            pan = row.get(4, '')        # Coluna E (índice 4)
            auth_text = row.get(28, '') # Coluna AC (índice 28)

            if not protocolo:
                continue

            codigo_aut = None
            # Verifica se a célula de autorização tem algum texto
            if auth_text and isinstance(auth_text, str):
                # Usa a nova expressão regular, mais robusta
                match = re.search(r'[Aa][Uu][Tt]:\s*(\d{6})', auth_text)
                if match:
                    codigo_aut = match.group(1)

            dados_planilha.append({
                'linha': i + 1,
                'protocolo': str(protocolo).strip(),
                'pan': str(pan).strip() if pan else 'N/A',
                'aut': codigo_aut if codigo_aut else 'N/A'
            })

        return dados_planilha

    except FileNotFoundError:
        print(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return []
    except Exception as e:
        print(f"ERRO: Ocorreu uma falha inesperada ao ler o arquivo Excel: {e}")
        return []


if __name__ == "__main__":
    print("--- Teste de Leitura de Planilha (Suporta .xlsx e .xlsb) ---")

    caminho = input("Por favor, insira o caminho completo para o arquivo Excel: ")
    aba = input("Por favor, insira o nome da planilha (ex: Planilha1): ")

    if not os.path.exists(caminho):
        print(f"\nERRO: O caminho '{caminho}' não existe. Verifique e tente novamente.")
    else:
        dados_extraidos = ler_dados_excel_avancado(caminho, aba)

        if dados_extraidos:
            print("\n--- Dados Extraídos com Sucesso ---\n")
            for item in dados_extraidos:
                print(
                    f"Linha: {item['linha']:<5} | Protocolo: {item['protocolo']:<15} | PAN: {item['pan']:<20} | Aut: {item['aut']}")
            print(f"\nTotal de {len(dados_extraidos)} registros encontrados.")
        else:
            print("\nNenhum dado foi extraído. Verifique se o arquivo e o nome da aba estão corretos.")

    input("\nPressione Enter para sair...")
