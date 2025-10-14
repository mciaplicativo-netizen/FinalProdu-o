# Produção & Estoque - Streamlit (Sincronizado com SQLite & Excel)

Este projeto implementa um app Streamlit que sincroniza dados entre um arquivo Excel e um banco SQLite local.
- Ao iniciar, o app pode importar dados do Excel para o DB.
- Todas as alterações no app gravam no SQLite e regravem as sheets correspondentes no Excel.

**Notas importantes**
- Sempre mantenha backup do Excel antes de usar pela primeira vez.
- Em ambientes com múltiplos usuários simultâneos, pode haver condições de corrida. Este app usa um lock simples de arquivo para reduzir conflitos.
- Ajuste nomes das sheets no sidebar conforme necessário.
