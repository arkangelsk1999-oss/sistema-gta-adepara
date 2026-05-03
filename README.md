# Sistema de Consulta GTA — ADEPARÁ
Desenvolvido por **Arkangelsk**

## Acesso inicial (Founder)
- E-mail: `founder@arkangelsk.com`
- Senha: `Arkangelsk@2025`
⚠️ Altere a senha no primeiro acesso.

## Estrutura
- `app.py` — aplicação principal Flask
- `database.py` — banco de dados e lógica de importação
- `relatorio.py` — geração de Excel e PDF
- `templates/` — páginas HTML
- `static/img/` — logos e brasão

## Instalação local
```bash
pip install -r requirements.txt
python app.py
```
Acesse: http://localhost:5000

## Deploy Railway
1. Conecte este repositório no Railway
2. Configure variável de ambiente: `SECRET_KEY=sua-chave-secreta`
3. Deploy automático
