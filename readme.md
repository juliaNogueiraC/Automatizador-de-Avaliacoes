# Automatizador de Avaliações

## Descrição
Este projeto foi criado para automatizar o processo de preenchimento de avaliações individuais para funcionários. Anteriormente, o preenchimento era feito manualmente, onde cada nome de funcionário precisava ser inserido individualmente em cada folha de avaliação. Com este script, é possível automatizar esse processo, economizando tempo e minimizando erros humanos.

## Funcionalidades
- Carrega uma lista de funcionários de um arquivo de texto.
- Preenche automaticamente o nome de cada funcionário em uma folha de avaliação.
- Gera uma avaliação personalizada em formato PDF para cada funcionário.
- Cria uma pasta contendo as avaliações personalizadas de cada funcionário.

## Uso
1. Coloque o modelo de folha de avaliação em formato Excel na pasta `modelo_avaliacao.xlsx`.
2. Crie um arquivo de texto chamado `lista_alunos.txt` contendo os nomes dos funcionários, um por linha.
3. Execute o script `gerar_avaliacoes.py`.
4. As avaliações personalizadas de cada funcionário serão geradas na pasta `avaliacoes_pdf`.

## Requisitos
- Python 3.x
- Bibliotecas Python: `openpyxl`, `reportlab` (ou `fpdf`)

## Instalação de Dependências
```bash
pip install openpyxl reportlab
