# Sistema GTA (Histórico de GTA)

Destino previsto: `CLAUDE.md` na pasta do repositório `sistema-gta-adepara`. Variante SQLite da skill `padrao-flask-adepara`.

- Flask + SQLite (FTS5 para busca). Bancos: `gta.db` (~28 GB, 2012–2025), `gta2026.db`, `gtv.db`, `usuarios.db`.
- Nunca copiar, abrir inteiro ou alterar `gta.db`. É banco de leitura após a importação.
- Excel com `xlsxwriter`, não `openpyxl`. Acima de 50.000 registros, ofereça ZIP.
- LAI e LGPD: CPF (11 dígitos) vira `***ANONIMIZADO***` nas exportações; CNPJ (14 dígitos) permanece.
- O banco de 2024 não tem `uf_orig`; trate `NULL` e mantenha a nota de aviso no Excel da LAI.
- Importação de planilhas não tem desfazer. Peça backup e confirmação antes de importar.
- Módulos de banco em `database_*.py`; rotas em `app.py`.
- Não há registro formal de alterações. Criar `docs/alteracoes.md` e `docs/apontamentos-manual.md` na primeira entrega.
- Pendências conhecidas: logo novo da ADEPARÁ, seção separada de 2026 na LAI, busca por número no GTV.
