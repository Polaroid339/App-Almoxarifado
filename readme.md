# 🗃️ Sistema de Almoxarifado com Interface Gráfica

Este projeto é uma aplicação de controle de almoxarifado desenvolvida em Python com interface gráfica Tkinter. Ele permite o gerenciamento de estoque, entradas e saídas de produtos, controle de EPIs, login de usuários e geração de relatórios, tudo com manipulação de dados em arquivos .csv.

---

## 📦 Funcionalidades

- Cadastro de produtos no estoque  
- Registro de entrada e saída de produtos  
- Cadastro e retirada de EPIs  
- Geração de relatórios em Excel e .txt  
- Interface gráfica amigável com abas e botões  
- Backup automático a cada 3 horas  
- Tela de login com verificação de operador  
- Correção automática de planilhas mal formatadas

---

## 🧰 Tecnologias Utilizadas

- Python 3.x  
- Tkinter  
- Pandas  
- Pandastable  
- CSV  
- PyInstaller (para empacotamento em .exe)

---

## 📁 Estrutura de Pastas

- main.py: Arquivo principal do sistema  
- Planilhas/: Armazena os arquivos .csv de estoque, entrada, saída e EPIs  
- Colaboradores/: Contém os arquivos de registro por colaborador  
- Relatorios/: Saída dos relatórios gerados  
- Backups/: Cópias de segurança automáticas dos arquivos  
- usuarios.py: Dicionário com usuários e senhas

---

## ▶️ Como Executar o Projeto

1. Verifique se o Python 3 está instalado na sua máquina.

2. Instale os pacotes necessários usando o seguinte comando no terminal:

```bash  
pip install pandas pandastable  
```

3. Execute o sistema com o comando:

```bash  
python main.py  
```

A interface gráfica será carregada com a tela de login.

---

## 🔐 Sistema de Login

Os usuários são definidos no arquivo usuarios.py. Exemplo de estrutura:

```python  
usuarios = {
    "admin": {"senha": "1234", "id": "001"},
    "usuario": {"senha": "senha123", "id": "002"}
}
```

Ao realizar login com sucesso, o sistema libera o acesso completo às funções.

---

## 🛠️ Gerar Executável .exe

Se desejar empacotar a aplicação em um executável para Windows, use o PyInstaller com o seguinte comando:

```bash  
python -m PyInstaller --onefile --name=Almoxarifado --windowed --add-data "Planilhas;Planilhas" main.py  
```

Este comando gera um .exe na pasta dist.

---

## 📤 Relatórios Exportáveis

A opção "Exportar" gera:

- Um arquivo Excel com todas as planilhas (Estoque, Entrada, Saída, EPIs)  
- Um arquivo .txt listando produtos esgotados (com quantidade igual a 0)

Os relatórios são salvos na pasta Relatorios/.

---

## 💾 Backup Automático

O sistema realiza backups automáticos das planilhas a cada 3 horas e armazena na pasta Backups/. Backups com mais de 3 dias são removidos automaticamente.

---

## 🧤 Controle de EPIs

- Cadastro de EPI com CA ou descrição  
- Atualização de quantidade se o EPI já existir  
- Registro de retiradas por colaborador  
- Arquivo gerado por colaborador e por mês (em Colaboradores/NOME/mes.csv)

---

## 📝 Observações Finais

- O sistema verifica se as planilhas estão corretamente formatadas ao iniciar.  
- O campo "VALOR TOTAL" é calculado automaticamente com base no valor unitário e na quantidade.  
- Todas as alterações feitas na tabela podem ser salvas com um clique no botão "Salvar Alterações".

---

## 👨‍💻 Autor

Desenvolvido por Victor Oliveira.  
Contato para dúvidas ou sugestões: github.com/Polaroid339
