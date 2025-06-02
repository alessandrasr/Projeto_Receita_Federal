# Consulta CPF Automatizada – Receita Federal  
Automação em **Python + Selenium** que lê uma lista de CPFs (+ data de nascimento) a partir de um arquivo **Excel**, consulta a situação cadastral de cada CPF no serviço público da Receita Federal e grava os nomes retornados em um novo arquivo Excel.

> **⚠️ Atenção:** CPF é dado sensível. Use esta ferramenta apenas para fins legítimos, com autorização dos titulares e em conformidade com a LGPD.

---

Funcionalidades
Descrição
✔️ Chrome em modo depuração	Abre o navegador já autenticado com seu perfil.
✔️ Detecção & clique no hCaptcha	Tenta resolver o checkbox inicial duas vezes antes de aguardar intervenção manual.
✔️ Leitura inteligente de Excel	Formata CPF com 11 dígitos, converte datas ao padrão exigido pelo site.
✔️ Registro de resultados	Gera Resultados_receita.xlsx pronto para colar na planilha ANALISAR.
✔️ Tratamento de erros	Mensagens claras no console e retomada automática da consulta linha a linha.
