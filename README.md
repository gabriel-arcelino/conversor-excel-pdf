# Conversor de Excel para PDF

Este projeto em Python tem como objetivo automatizar a conversão de arquivos Excel (.xlsx ou .xls) para PDF, preservando a formatação original das planilhas, e replicando a estrutura de pastas no diretório de destino.

## Funcionalidades

- Converte arquivos Excel para PDF utilizando o próprio Excel (via `pywin32`).
- Mantém a estrutura de pastas do diretório de origem.
- Gera um PDF por arquivo Excel, preservando o layout, gráficos, e formatação.
- Suporta múltiplas planilhas em um único arquivo Excel.

## Pré-requisitos

- **Sistema Operacional**: Windows (necessário para utilizar o Excel via automação).
- **Microsoft Excel**: Deve estar instalado no sistema.
- **Python 3.x**: Certifique-se de ter o Python instalado e configurado.
- **Bibliotecas necessárias**:
  - `pywin32`: Biblioteca para controlar o Excel via automação do Windows.

### Instalação das dependências

Para instalar a biblioteca necessária, utilize o seguinte comando:

```bash
pip install pywin32
```

## Como usar

1. **Clone o repositório**:

   ```bash
   git clone https://github.com/seu-usuario/seu-repositorio.git
   cd seu-repositorio
   ```

2. **Configure os diretórios**:

   Edite o código para definir os caminhos corretos para os arquivos Excel e o destino dos PDFs.

   - `diretorio_base`: Caminho para o diretório contendo os arquivos Excel.
   - `diretorio_destino_base`: Caminho onde os PDFs serão salvos.

3. **Execute o script**:

   Após configurar os diretórios, execute o script principal:

   ```bash
   python app.py
   ```

4. **Resultado**:

   O script irá:
   - Percorrer as pastas e subpastas em busca de arquivos Excel.
   - Converter cada arquivo para PDF.
   - Salvar os PDFs no diretório de destino, mantendo a mesma estrutura de pastas.

## Estrutura do projeto

```
projeto-conversor-excel-pdf/
│
├── app.py                # Script principal de conversão
├── README.md              # Este arquivo
├── requisitos.txt         # Bibliotecas Python necessárias
└── ...
```

## Contribuições

Fique à vontade para contribuir com melhorias, novas funcionalidades ou correções de bugs. Faça um fork do repositório, crie uma branch para suas alterações e envie um Pull Request.

---

Se precisar de mais algum ajuste, é só avisar!
