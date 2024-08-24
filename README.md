# Automação de Cadastro de Produtos

Este projeto automatiza o cadastro de produtos de uma planilha Excel em um sistema web. Ele usa bibliotecas Python para manipular a planilha e interagir com a interface do sistema via simulação de teclado e mouse.

## Objetivo

Automatizar o processo de preenchimento e cadastro de produtos no sistema web usando dados provenientes de uma planilha Excel.

## Tecnologias Utilizadas

- **Python**
- **OpenPyXL**: Manipulação de arquivos Excel.
- **PyAutoGUI**: Automação de interações com a interface do usuário (cliques e teclas).
- **Pyperclip**: Facilita a cópia de dados para a área de transferência, mantendo os caracteres especiais (acentos agudos, cinrunflexos, etc).
- **Time**: Para gerenciar delays nas ações automatizadas.

## Requisitos

- Python 3.x
- Bibliotecas Python: `openpyxl`, `pyautogui`, `pyperclip`
- A planilha deve estar no caminho `data/produtos_ficticios.xlsx` com os dados dos produtos na aba "Produtos".

## Estrutura da Planilha

A planilha deve conter os seguintes campos (colunas) na aba "Produtos", a partir da linha 2:

1. Nome do Produto
2. Descrição
3. Categoria
4. Código do Produto
5. Peso
6. Dimensões
7. Preço
8. Quantidade
9. Validade
10. Cor
11. Tamanho (Pequeno, Médio, Grande)
12. Material
13. Fabricante
14. País de Origem
15. Observações
16. Código de Barras
17. Localização no Armazém

## Instruções de Uso

1. Certifique-se de que o sistema web esteja aberto na página inicial de cadastro de produtos: [Cadastro de Produtos](https://cadastro-produtos-devaprender.netlify.app/index.html).
2. Execute o script Python para iniciar o processo de automação.
3. Quando o script iniciar não mexer no teclado ou mouse até o processo finalizar.
4. O script irá:
   - Ler os dados da planilha.
   - Preencher automaticamente os campos no sistema web.
   - Clicar nos botões necessários para concluir o cadastro.
   - Repetir o processo para cada produto na planilha.

## Executando o Script

1. Clone este repositório.
2. Instale as dependências necessárias:
   ```bash
   pip install openpyxl pyautogui pyperclip
   ```
3. Execute o script:
   ```bash
   python automacao_planilha.py
   ```

## Notas

- A automação simula cliques e teclas, então é importante que a interface do sistema web esteja visível e em foco durante a execução do script.
- Certifique-se de que as coordenadas dos cliques no código correspondem às posições dos campos no seu monitor.
- Pode ser necessário ajustar os delays (`sleep`) dependendo da velocidade de resposta do sistema.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests.

## Licença

Este projeto está licenciado sob a licença MIT.
```

Esse `README.md` reflete as funcionalidades e o uso do código informado.