# Gerenciador de Almoxarifado

Este aplicativo é um sistema de gerenciamento de almoxarifado e estoque desenvolvido em Python com interface gráfica utilizando Tkinter. Ele permite gerenciar o estoque, registrar entradas e saídas de produtos, e gerar relatórios.


## Funcionalidades

### 1. **Gerenciamento de Estoque**
- Visualize os produtos cadastrados no estoque.
- Pesquise produtos por código ou descrição.
- Atualize as informações do estoque.

### 2. **Cadastro de Produtos**

- Cadastre novos produtos no estoque com as seguintes informações:

  - Descrição
  - Quantidade
  - Valor Unitário
  - Localização
 
Código e Data de Cadastro serão gerados dinamicamente.

### 3. **Movimentação de Produtos**

- **Registrar Entrada**:
  - Adicione quantidades ao estoque de produtos existentes.
  - Registre a entrada com informações como código, quantidade e valor total.

- **Registrar Saída**:
  - Retire quantidades do estoque de produtos existentes.
  - Registre a saída com informações como código, quantidade e solicitante.

### 4. **Relatórios**

- Exporte os dados do estoque, entradas e saídas para um arquivo Excel.
- Gere um relatório de produtos esgotados em formato `.txt`.

### 5. **Filtros e Pesquisa**

- Pesquise produtos no estoque por qualquer termo.
- Limpe os filtros para visualizar todos os produtos novamente.

- **Planilhas/**: Contém os arquivos CSV que armazenam os dados do estoque, entradas e saídas.
- **app.py**: Código principal do sistema.
- **favicon.ico**: Ícone do aplicativo.
- **README.md**: Documentação do projeto.

## Requisitos

- Python 3.8 ou superior
- Bibliotecas Python:
  
  - `pandas`
  - `tkinter`
  - `csv`
  - `shutil`
  - `datetime`
  - `pandastable`

Para instalar as dependências, execute:
```bash
pip install -r requirements.txt
```

## Como Executar
1. Certifique-se de que a pasta Planilhas está no mesmo diretório que o arquivo app.py.
2. Execute o arquivo app.py:
```bash
python app.py
```

## Gerar Executável (.exe)

Para criar um executável do sistema, utilize o PyInstaller:

1. Instale o PyInstaller:
```bash
pip install pyinstaller
```

2. Gere o executável:
```bash
python -m PyInstaller --onefile --name=Almoxarifado --windowed --icon=favicon.ico --add-data "Planilhas;Planilhas" app.py
```

O executável será gerado na pasta `dist/` com o nome Almoxarifado.exe.

## Como Usar

### Aba Estoque
1. Visualize os produtos cadastrados no estoque.
2. Pesquise produtos utilizando o campo de busca.
3. Atualize ou salve alterações feitas na tabela.

### Aba Cadastro
1. Preencha os campos de descrição, quantidade, valor unitário e localização.
2. Clique no botão "Cadastrar" para adicionar o produto ao estoque.
3. Aba Movimentação

### Aba Movimentação

#### - Registrar Entrada:
1. Insira o código do produto e a quantidade a ser adicionada.
2. Clique em "Registrar Entrada".

#### - Registrar Saída:
1. Insira o código do produto, o nome do solicitante e a quantidade a ser retirada.
2. Clique em "Registrar Saída".

### Relatórios
1. Clique no botão "Exportar" na aba Estoque para gerar:
2. Um arquivo Excel com os dados do estoque, entradas e saídas.
3. Um arquivo .txt com os produtos esgotados.

## Melhorias Futuras
- Adicionar autenticação de usuários;
- Implementar alertas para estoque baixo;
- Adicionar suporte a múltiplos idiomas;
- Criar gráficos e relatórios visuais.

## Autor
Desenvolvido por Victor Oliveira (Polaroid339).

## Licença
Este projeto é licenciado sob a MIT License, mais informações em `LICENSE.TXT`
