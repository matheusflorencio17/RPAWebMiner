<<<<<<< HEAD
## RPA de Captura e Filtro de Produtos da Magazine Luiza

Este RPA automatiza a busca, captura e filtragem de produtos no site da Magazine Luiza, usando **Python** e **Selenium**. Ele verifica a disponibilidade do site, busca um item específico, captura dados sobre os produtos, filtra os resultados e envia um relatório por e-mail.

### Funcionalidades:

1. **Verificação de Disponibilidade**: Verifica se o site está disponível.
2. **Busca de Produto**: Pesquisa um item no site (nome configurado via `.env`).
3. **Captura de Dados**: Coleta informações como nome, quantidade de avaliações e URL dos produtos.
4. **Filtragem e Salvamento**: Filtra os produtos com base nas avaliações e salva em um arquivo Excel.
5. **Envio de E-mail**: Envia o arquivo Excel como anexo por e-mail.

### Tecnologias Utilizadas:

- **Python 3.x**
- **Selenium**: Automação de navegação na web.
- **Pandas**: Manipulação e salvamento de dados.
- **SMTP**: Envio de e-mails.

### Pré-Requisitos:

1. **Python 3.x** instalado.
2. Instalar dependências:

   ```bash
   pip install selenium pandas requests python-dotenv webdriver-manager
   ```

3. **Arquivo `.env`** com as seguintes variáveis:
   - `ITEM_NAME`: Nome do item a ser buscado.
   - `USER_EMAIL`: Seu e-mail.
   - `USER_PASSWORD`: Senha do seu e-mail.
   - `SEND_EMAIL`: E-mail do destinatário.

### Como Executar:

1. **Configure o arquivo `.env`** com suas credenciais e o nome do item.
2. **Execute o script**:

   ```bash
   python index.py
   ```

   O Index:
   - Verifica o site.
   - Busca o item.
   - Captura os produtos e os salva em um Excel.
   - Envia o relatório por e-mail.

### Estrutura de Diretórios:

```
.
├── .env                # Arquivo de variáveis de ambiente
├── script.py           # Código do RPA
├── log.txt             # Log de execução
└── Output/             # Pasta com o arquivo Excel 
    └── Notebooks.xlsx  # Relatório com os produtos
** Após a consulta é criado a pasta e o arquivo (Output - Notebooks.xlsx) através do codigo.
```

### Notas:

- **Limite de Páginas**: O script captura até 2 páginas de resultados (ajustável).
- **Lembrando**: Caso não queira definir limite, o Laço condicional irá se repetir até que não seja mais encontrado produtos no qual foi feita pesquisa.
- **Erro de Permissão**: Se houver erro ao salvar o arquivo, verifique permissões ou se o arquivo está aberto.

### Contribuição:

Caso encontre algum problema ou tenha sugestões, sinta-se à vontade para abrir uma **issue** ou contribuir para o código!

=======
# RPAWebMiner
>>>>>>> b9a69e118a2ea918b8091c3b4c0a1c35f96b1482
