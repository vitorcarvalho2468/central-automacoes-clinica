#!/usr/bin/env bash
# Recria o repositorio git e sobe pro GitHub. Rode uma vez, dentro da pasta do projeto:
#     bash publicar.sh
set -e
cd "$(dirname "$0")"

echo ">>> Preparando o .gitignore..."
cat > .gitignore <<'EOF'
# Python
__pycache__/
*.py[cod]
*.egg-info/
.venv/
venv/
env/

# PyInstaller
build/
dist/
*.manifest

# Editores
.idea/
.vscode/
*.swp
.DS_Store

# Planilhas: NUNCA comitar dados reais de paciente.
# Ignora todo .xlsx/.xlsm, menos os exemplos ficticios de sample_data/
*.xlsx
*.xlsm
!sample_data/**

# Perfis de navegador / dados locais dos robos (ficam em ~/CentralAutomacoes)
CentralAutomacoes/
EOF

echo ">>> Iniciando o repositorio..."
git init -b main >/dev/null 2>&1 || git init >/dev/null 2>&1
git checkout -b main >/dev/null 2>&1 || true
git config user.name "Vitor Carvalho"
git config user.email "vitorcarvalho2468@users.noreply.github.com"
git config commit.gpgsign false

add_commit () {
  msg="$1"; shift
  git add "$@" 2>/dev/null || true
  git commit -q -m "$msg" 2>/dev/null || true
}

echo ">>> Criando o historico de commits..."
add_commit "Base do projeto: leitura tolerante de .xlsx, navegador e utilidades" \
  .gitignore LICENSE requirements.txt xlsx_seguro.py navegador.py feriados_br.py
add_commit "Automacao de feriado (redistribuicao de agenda) e extrator de TR" \
  feriado_engine.py gerador_agenda.py automacao_anterior.py
add_commit "Robos dos portais (SulAmerica, Porto, Unimed) com Playwright + conferencia de faturamento" \
  robo_sulamerica.py robo_porto.py robo_unimed.py guia_porto_pdf.py conferencia_faturamento.py \
  interface_robo.py interface_porto.py interface_unimed.py
add_commit "Uniao de conferencias: junta as planilhas da semana numa geral padronizada" \
  uniao.py interface_uniao.py
git add -A
git commit -q -m "App principal, scripts de build, documentacao e dados de exemplo ficticios" 2>/dev/null || true

echo ">>> Conectando no GitHub..."
git remote remove origin 2>/dev/null || true
git remote add origin https://github.com/vitorcarvalho2468/central-automacoes-clinica.git
git branch -M main

echo ""
echo ">>> Tudo pronto localmente. Historico:"
git log --oneline
echo ""
echo ">>> Subindo pro GitHub (pode abrir o navegador pra voce logar)..."
git push -u origin main
echo ""
echo ">>> Pronto! Abra: https://github.com/vitorcarvalho2468/central-automacoes-clinica"
