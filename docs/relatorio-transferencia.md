# Relatório de Transferência — Sistema Histórico de GTA
**ADEPARÁ — Agência de Defesa Agropecuária do Estado do Pará**
Data: novembro/2026 | Autor: Gilliard (Arkangelsk)

---

## 1. Visão geral

**Nome do sistema:** Sistema GTA — Consulta ao Histórico de Emissões de GTA

**Objetivo:** Permitir que servidores da ADEPARÁ consultem o histórico completo de emissões de Guias de Trânsito Animal (GTA) do Estado do Pará, com exportação de dados para Excel/CSV e geração de relatórios conforme a Lei de Acesso à Informação (LAI), com anonimização automática conforme a LGPD.

**Quem usa:** Servidores da ADEPARÁ, principalmente do setor GTAGRO (Gerência de Trânsito Agropecuário). Acesso por rede interna (IP local). Um colega da área comum também acessa via launcher dedicado.

**Fase:** Sistema em produção, em uso interno. Recebe atualizações e melhorias incrementais. Há módulos adicionais (GTA 2026 e GTV) já funcionando. Algumas funcionalidades planejadas ainda não implementadas (ver seção 8).

**GitHub:** `arkangelsk1999-oss/sistema-gta-adepara`

---

## 2. Tecnologia

| Item | Detalhe |
|------|---------|
| Linguagem | Python 3.11 |
| Framework web | Flask |
| Banco de dados | SQLite (sem versão específica registrada; nativo do Python) |
| Busca textual | SQLite FTS5 (módulo nativo) |
| Geração de Excel | xlsxwriter (substituiu openpyxl por ser até 10× mais rápido em arquivos grandes) |
| Geração de PDF | reportlab |
| Frontend | HTML + CSS + JavaScript puro (sem framework JS) |
| Autenticação | Sessão Flask com cookie + reCAPTCHA v3 (Google) |
| Como roda | Local, Windows 10, direto pelo CMD com `python app.py` ou via arquivo `.bat` |
| Docker | Não utilizado |
| Servidor de produção | Não definido (roda em modo desenvolvimento Flask; instalação em Linux está planejada mas não realizada) |

---

## 3. Estrutura

### Pastas e arquivos principais

```
C:\Users\57216615\sistema-gta-adepara\   ← pasta principal (PC de trabalho)
C:\sistema-gta-adepara\                  ← cópia para área comum (colegas)

├── app.py                    # Ponto de entrada — rotas Flask de todos os módulos
├── database_gta.py           # Funções de acesso ao banco GTA 2012–2025 (gta.db)
├── database_gta2026.py       # Funções de acesso ao banco GTA 2026+ (gta2026.db)
├── database_gtv.py           # Funções de acesso ao banco GTV (gtv.db)
├── database_usuarios.py      # Autenticação e gestão de usuários
├── gta.db                    # Banco principal (~28 GB, 6,77 M registros, 2012–2025)
├── gta2026.db                # Banco GTA 2026+ (~239 K registros)
├── gtv.db                    # Banco GTV (~102 K registros)
├── usuarios.db               # Banco de usuários e logs de acesso
├── requirements.txt          # Dependências Python
├── venv\                     # Ambiente virtual (NÃO versionado no git)
├── templates\
│   ├── base.html             # Layout base (navbar, estilos globais)
│   ├── login.html            # Tela de login com reCAPTCHA v3
│   ├── index.html            # Consulta GTA 2012–2025 (módulo principal)
│   ├── gta2026.html          # Consulta GTA 2026+
│   ├── gtv.html              # Consulta GTV
│   ├── importar.html         # Tela de importação de planilhas (roteamento por nome de arquivo)
│   └── ...                   # Outras telas (LAI, gerenciamento de usuários, etc.)
└── static\
    ├── adepara_logo.png      # Logo da ADEPARÁ (substituição pendente — ver seção 8)
    └── ...                   # CSS, JS e outros assets
```

### Tabelas do banco

#### `gta` (em `gta.db`) — tabela principal

| Coluna | Tipo | Observação |
|--------|------|------------|
| rowid | INTEGER | Chave primária implícita do SQLite |
| num_gta | TEXT | Número da GTA (ex.: PA672222O) |
| nome_rem | TEXT | Nome do remetente |
| cpf_cnpj_rem | TEXT | CPF (11 dígitos) ou CNPJ (14 dígitos) do remetente |
| nome_dest | TEXT | Nome do destinatário |
| cpf_cnpj_dest | TEXT | CPF/CNPJ do destinatário |
| emissor | TEXT | Usuário do SIGEAGRO que emitiu a GTA |
| finalidade | TEXT | Finalidade do transporte (ABATE, RECRIA, etc.) |
| especie | TEXT | Espécie animal |
| qtd_animais | INTEGER | Quantidade de animais |
| municipio_orig | TEXT | Município de origem |
| uf_orig | TEXT | UF de origem (ausente em 2024 — ver seção 8) |
| municipio_dest | TEXT | Município de destino |
| uf_dest | TEXT | UF de destino |
| data_emissao | TEXT | Data de emissão (formato ISO: YYYY-MM-DD) |
| ano | INTEGER | Ano extraído para particionamento lógico |
| ... | ... | Outras colunas variam por grupo de importação (A/B/C/D) |

> **Tabela virtual FTS5:** `gta_fts` — indexa `nome_rem`, `nome_dest`, `cpf_cnpj_rem`, `cpf_cnpj_dest`, `emissor`. É o que torna as buscas rápidas em 6,77 M registros.

#### `gta2026` (em `gta2026.db`)
Mesma estrutura lógica; schema do Grupo D (2025+). Possui FTS5 próprio.

#### `gtv` (em `gtv.db`)
Estrutura análoga para Guias de Trânsito Vegetal.

#### `usuarios` (em `usuarios.db`)

| Coluna | Tipo | Observação |
|--------|------|------------|
| id | INTEGER | Chave primária |
| login | TEXT | Nome de usuário |
| senha_hash | TEXT | Hash da senha (nunca texto puro) |
| nivel | TEXT | `founder` / `master` / `normal` |
| ativo | INTEGER | 1 = ativo, 0 = bloqueado |

#### `logs_acesso` (em `usuarios.db`)

| Coluna | Observação |
|--------|------------|
| id | Chave primária |
| usuario_id | FK para `usuarios` |
| ip | IP de origem da requisição |
| localidade | Localidade inferida do IP |
| data_hora | Timestamp da ação |
| acao | Descrição da ação realizada |

---

## 4. Padrões que adotamos

- **Idioma dos nomes de variáveis/funções:** português (ex.: `buscar_gtas()`, `listar_finalidades_2026()`)
- **Idioma dos comentários no código:** português
- **Formato de datas:** ISO 8601 (`YYYY-MM-DD`) no banco; exibição formatada no frontend
- **Nomes de tabelas:** snake_case em português (`gta`, `gta2026`, `gta_fts`, `logs_acesso`)
- **Nomes de arquivos de banco:** `gta.db`, `gta2026.db`, `gtv.db`, `usuarios.db`
- **Mensagens de erro no frontend:** texto direto em português, prefixadas com ❌ no JS; erros positivos com ✅
- **Separação de módulos:** cada módulo de banco de dados tem seu próprio arquivo `database_*.py`; o `app.py` centraliza as rotas
- **Caching:** listas de finalidades são cacheadas em memória (`listar_finalidades()` / `listar_finalidades_2026()`) para evitar queries repetidas
- **Exportações grandes:** usam `xlsxwriter` (não `openpyxl`) e, para volumes > 50.000 registros, oferecem ZIP compactado
- **LGPD:** CPF (11 dígitos) → `***ANONIMIZADO***` nas exportações LAI; CNPJ (14 dígitos, pessoa jurídica/pública) → mantido

---

## 5. Acessos e segurança

### Perfis de usuário

| Nível | Permissões |
|-------|-----------|
| `founder` | Acesso total; invisível para outros usuários; pode exportar CSV e ZIP |
| `master` | Consulta + exportação Excel + CSV + ZIP; pode ver logs |
| `normal` | Consulta + exportação Excel apenas |

### Autenticação
- Formulário de login com usuário + senha
- Sessão Flask com cookie (timeout de inatividade configurado)
- reCAPTCHA v3 invisível no login (chave: [OCULTO]; ativado via `RECAPTCHA_ATIVO=true` no ambiente)
- Sem fluxo automatizado de reset de senha — recuperação por canal humano

### Logs
- Toda ação de acesso registrada em `logs_acesso`: usuário, IP, localidade, timestamp e ação
- IP mapeado para localidade (município/estado) para auditoria

### Validações de entrada
- CPF/CNPJ: somente dígitos, validação de tamanho
- Campos de texto: convertidos para maiúsculas no frontend antes do envio
- Número de GTA: somente dígitos no campo de busca

### Proteções
- Não há exclusão de registros do banco de GTA (banco é somente leitura após importação)
- Importação de planilhas restrita a usuários autenticados
- reCAPTCHA bloqueia automação no login em produção

---

## 6. Minhas preferências e correções

### Regras absolutas (nunca violar)

1. **Sempre reescrever o arquivo completo** — nunca mandar só o trecho alterado, nunca `find-and-replace`, nunca `sed`. O arquivo inteiro vai no VS Code.
2. **Sempre informar o caminho completo do arquivo** antes de pedir para editar.
3. **Sempre especificar qual terminal pelo número** (Terminal 1, Terminal 2, etc.).
4. **Sempre abrir um terminal novo (+)** — nunca reaproveitar terminal que já está rodando algo.
5. **Usar CMD, não PowerShell.**
6. **Nunca mandar comandos Python `-c` de múltiplas linhas** para o terminal — isso causa problemas.
7. **Arquivos grandes → gerar como download**, não colar no chat.
8. **Máximo 3 passos por vez** — não despeja tudo de uma vez.
9. **Comunicação direta, sem enrolação** — sem "claro!", "ótima pergunta!", "vamos lá!".
10. **Commit e push ao final de cada sessão de trabalho.**

### Correções que fiz ao longo do projeto

- **Período do banco:** corrigi a IA que dizia "2010 a 2025" — o banco é de fato de **2012 a 2025** (os registros de 2010/2011 eram resíduos que foram removidos/ignorados).
- **Zeros à esquerda em CPF:** bug crítico detectado por mim — 303.249 registros dos anos 2024/2025 estavam sem zeros à esquerda no CPF; exigi a correção.
- **Coluna `uf_orig` ausente em 2024:** identifiquei que o ano de 2024 não exporta essa coluna no SIGEAGRO; exigi nota de aviso amarela no Excel da LAI.
- **xlsxwriter vs openpyxl:** substituição foi minha decisão após lentidão nos testes de exportação LAI.
- **A IA gerou um `index.html` sem eu pedir** — reclamei diretamente: *"Mano! O que diabos fez?"*. A IA não deve gerar código que não foi pedido.
- **A IA tentou acessar o sistema via browser sem autorização** — reclamei. Nunca fazer isso.

### Meu jeito de trabalhar

- Sou **Fiscal Estadual Agropecuário e Médico Veterinário** — tenho domínio do negócio, não só do código. Minhas decisões sobre regras de negócio são finais.
- Metodologia: **humano como arquiteto, IA como executor**. Eu defino o que e por quê; a IA faz o como.
- Prefiro **trabalhar em passos pequenos e validar cada um** antes de seguir.
- Me irrita: enrolação, perguntas óbvias, código gerado sem ser pedido, suposições erradas sobre o negócio, e a IA inventar soluções sem perguntar.
- Tom esperado: direto, técnico, sem bajulação.

---

## 7. Decisões importantes

| Decisão | Motivo |
|---------|--------|
| **SQLite em vez de PostgreSQL/MySQL** | Banco de leitura com tamanho fixo — sem inserções frequentes em produção; simplicidade de deploy |
| **FTS5 para busca textual** | Performance em 6,77 M registros — busca por nome/CPF era inviável com `LIKE` |
| **xlsxwriter em vez de openpyxl** | Até 10× mais rápido na geração dos Excels da LAI; openpyxl travava para bases grandes |
| **Quatro grupos de importação (A/B/C/D)** | O SIGEAGRO mudou o schema de exportação em 2022, 2024 e 2025 — sem padronização não é possível importar todas as eras |
| **Importação por tela `/importar` com roteamento por nome de arquivo** | Facilita para o usuário não técnico — sem CLI, sem scripts avulsos |
| **CPF → `***ANONIMIZADO***` na LAI; CNPJ mantido** | LGPD: CPF é dado de pessoa física (sigiloso); CNPJ é de pessoa jurídica/entidade pública (dado público) |
| **reCAPTCHA desativável via variável de ambiente** | Facilita desenvolvimento local sem bloquear o login; em produção liga com `RECAPTCHA_ATIVO=true` |
| **Módulo GTA 2026 separado (`gta2026.db`)** | Schema diferente do banco histórico; banco principal não recebe mais inserções — separar evita migrações no banco de 28 GB |
| **Perfil `founder` invisível para outros usuários** | Padrão de segurança adotado em todos os sistemas do Gilliard — suporte sem exposição institucional |
| **Recuperação de senha por canal humano** | Elimina vetor de ataque automatizado (reset por e-mail) em sistema sensível de órgão público |
| **Dados ficam na ADEPARÁ (não em nuvem)** | Conformidade com LGPD — dados de produtores rurais não saem da infraestrutura do órgão |

---

## 8. Pendências e problemas conhecidos

### Funcionalidades planejadas e não implementadas

| Item | Descrição |
|------|-----------|
| **Instalação em servidor Linux** | Sistema roda hoje só em Windows; migração para Linux planejada, não realizada |
| **reCAPTCHA em produção** | Requer domínio registrado e `RECAPTCHA_ATIVO=true`; hoje desativado em desenvolvimento |
| **Novo logo da ADEPARÁ** | Arquivo `adepara_logo.png` precisa ser substituído pelo logo atual |
| **LAI 2026 — seção separada** | Meses de 2026 precisam de seção separada no relatório LAI (não misturar com 2012–2025) |
| **Módulo GTV — busca por número de GTV** | Analogamente ao `num_gta` implementado no GTA, o GTV ainda não tem esse campo de busca |
| **PROJUR — novo sistema** | Banco para controle de processos judiciais — novo projeto, ainda não iniciado (conversa separada planejada) |

### Dívidas técnicas e bugs conhecidos

| Item | Descrição |
|------|-----------|
| **CPFs sem zeros à esquerda (pré-correção)** | Ambientes que não receberam o fix de outubro/2025 podem ter 303.249 registros com CPF incorreto |
| **Schema heterogêneo por período** | Colunas ausentes em certos anos (ex.: 2024 não tem `uf_orig`, `qtd_total`, `meio_transporte`); qualquer nova feature que use essas colunas precisa tratar o `NULL` |
| **Sem rollback de importação** | Se uma planilha for importada com erro, não há desfazer automático |
| **SQLite sem escrita concorrente** | Importações longas podem travar momentaneamente o acesso para outros usuários (leitura é concorrente, escrita não) |
| **Sem API autenticada por token** | Toda automação externa precisa simular o login por formulário e manter cookie de sessão |

---

## 9. Como registrávamos as alterações

Não havia um arquivo `CHANGELOG.md` formal mantido neste projeto durante esta conversa. As decisões e histórico foram registrados principalmente:

- Nas conversas com a IA (este chat e chats anteriores do projeto)
- Em um **bloco de contexto ("Resumo do Dia")** que o Gilliard colava no início de novas conversas, com o estado atual do projeto, pendências e preferências de trabalho
- O GitHub (`arkangelsk1999-oss/sistema-gta-adepara`) deveria ter o histórico de commits, mas o estado atual do repositório (se está atualizado) é **não definido** nesta sessão

Formato do bloco de contexto usado pelo Gilliard (reconstituído):
```
═══════════════════════════════════════════════════
SISTEMA GTA — CONTEXTO ATUAL
[data]
ESTADO: [descrição do que está funcionando]
PENDÊNCIAS: [lista]
PREFERÊNCIAS DE TRABALHO: [lista das regras]
TRABALHO ACADÊMICO: [metodologia, leis aplicáveis]
═══════════════════════════════════════════════════
```

---

## 10. Dez regras para a próxima IA trabalhar com Gilliard neste projeto

1. **Sempre reescreva o arquivo inteiro** — nunca envie trecho, diff, ou instrução de "substitua a linha X". O arquivo completo vai para o VS Code.
2. **Informe o caminho completo do arquivo** antes de qualquer edição (`C:\Users\57216615\sistema-gta-adepara\templates\index.html`, por exemplo).
3. **Especifique o terminal pelo número** ("Terminal 1", "Terminal 2") e sempre mande abrir um novo (+) em vez de reutilizar um ocupado.
4. **Use CMD, não PowerShell.** Se o usuário estiver no PowerShell, mande ele mudar.
5. **Máximo 3 passos por vez.** Aguarde confirmação antes de continuar.
6. **Nunca gere código que não foi pedido.** Não "antecipe" implementações, não refatore por conta própria, não crie arquivos extras sem ordem explícita.
7. **Nunca tente acessar o sistema pelo browser, fazer login, testar rotas ou executar comandos no servidor sem autorização explícita.**
8. **Arquivos grandes → gere como download** (arquivo `.py`, `.html`, `.md`) em vez de colar no chat.
9. **Seja direto.** Sem "ótima pergunta!", sem enrolação, sem listar 10 opções quando ele pediu uma solução. Se não souber, diga "não sei" e pergunte o que falta.
10. **Ao final de cada sessão de trabalho, lembre de fazer commit e push** no GitHub (`arkangelsk1999-oss/sistema-gta-adepara`), mesmo que seja um commit pequeno.
