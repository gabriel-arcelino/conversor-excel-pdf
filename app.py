import os
import win32com.client as win32

def encontrar_arquivos_excel(diretorio):
    arquivos_excel = []
    for root, dirs, files in os.walk(diretorio):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                arquivos_excel.append(os.path.join(root, file))
    return arquivos_excel

def criar_estrutura_pastas(destino):
    if not os.path.exists(destino):
        os.makedirs(destino)

def converter_excel_para_pdf(arquivo_excel, nome_pdf):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False  # Excel será executado em segundo plano
    workbook = None  

    try:
        # Tenta abrir o arquivo no modo de leitura
        workbook = excel.Workbooks.Open(arquivo_excel, ReadOnly=1)
        
        # Exportar para PDF (parametro 0 = PDF)
        workbook.ExportAsFixedFormat(0, nome_pdf)

        print(f"PDF gerado: {nome_pdf}")
    except Exception as e:
        print(f"Erro ao processar {arquivo_excel}: {e}")
    finally:
        # Somente fecha o workbook se ele foi aberto com sucesso
        if workbook:
            workbook.Close(SaveChanges=0)
        excel.Quit()

def processar_arquivos_excel(diretorio_base, diretorio_destino_base):
    arquivos_excel = encontrar_arquivos_excel(diretorio_base)

    for arquivo in arquivos_excel:
        # Obter o caminho relativo do arquivo Excel em relação ao diretório base
        caminho_relativo = os.path.relpath(arquivo, diretorio_base)

        # Substituir extensão do arquivo para '.pdf' e manter o mesmo nome
        caminho_relativo_pdf = os.path.splitext(caminho_relativo)[0] + '.pdf'

        # Caminho completo do diretório de destino para o PDF
        caminho_completo_pdf = os.path.join(diretorio_destino_base, caminho_relativo_pdf)

        # Criar a estrutura de pastas no diretório de destino, se necessário
        criar_estrutura_pastas(os.path.dirname(caminho_completo_pdf))

        # Converter o arquivo Excel para PDF
        converter_excel_para_pdf(arquivo, caminho_completo_pdf)

# Defina o caminho inicial para procurar os arquivos Excel e salvar os PDFs
diretorio_base = r'C:\Users\User\Documents\MAPAS 2024 - 2º BIMESTRE'
diretorio_destino_base = r'C:\Users\User\Documents\Gabriel Arcelino\Projetos\Converter Excel PARA PDF\MAPAS  2024 - 2º BIMESTRE-PDF'

processar_arquivos_excel(diretorio_base, diretorio_destino_base)
