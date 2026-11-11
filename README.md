# Central de Automações da Clínica

Um aplicativo de desktop que automatiza as tarefas chatas e repetitivas do dia a dia
de uma clínica de terapias: conferência de guias nos portais dos convênios, organização
de agendas e consolidação de planilhas.

Comecei esse projeto pra resolver um problema meu de verdade. Toda semana eu perdia horas
abrindo guia por guia nos portais dos convênios pra conferir se o atendimento tinha sido
confirmado, remanejando agenda quando caía feriado no meio da semana, e no fim juntando um
monte de planilha numa só na mão. Fui automatizando cada pedaço disso ao longo dos últimos meses, no uso real da clínica, e com
o tempo virou essa "central" com um botão pra cada tarefa.

> **Importante:** o programa nunca guarda login ou senha. Em todos os robôs que acessam portal,
> **o login é feito por você, na mão**, na janela do navegador que abre — o robô só continua
> a partir daí. Nenhum dado de paciente é armazenado pelo app.

---

## O que ela faz

Cada automação tem uma janela própria, abertas pela tela inicial:

| Automação | O que resolve |
|---|---|
| **Extrator de TR** | Limpa a coluna "Observação" das agendas, deixando só o código da guia (TR...). |
| **Redistribuição de feriado** | Pega a agenda da semana e redistribui os atendimentos do dia do feriado pelos outros dias, respeitando limites por paciente/unidade. Descobre o feriado sozinho. |
| **Robô SulAmérica / Porto / Unimed** | Abrem o portal do convênio (você faz o login), conferem a leitura/confirmação de cada guia e devolvem a planilha com o status. Usam um navegador real (Playwright) com a sua sessão. |
| **União de conferências** | Junta todas as conferências da semana (uma por convênio/dia) numa única "geral" padronizada de 12 colunas. |
| **Limpar prints de erro** | Apaga os prints/HTML que os robôs salvam pra depuração, pra não acumular lixo no disco. |

### Telas

> _Coloque aqui um print da tela inicial e um GIF de alguma automação rodando._
> _Salve as imagens em `docs/` e referencie, por exemplo:_
>
> `![Tela inicial](docs/tela-inicial.png)`

### Alguns detalhes de que eu me orgulho

- **Leitura tolerante de planilhas** (`xlsx_seguro.py`): o sistema da clínica exporta `.xlsx`
  com uma formatação de cor levemente fora do padrão que o Excel abre mas o `openpyxl` quebra.
  Em vez de pedir pra todo mundo "salvar de novo", o código conserta a formatação **em memória**
  (o arquivo original não é tocado) e segue a vida.
- **Padronização automática na união**: cada convênio exporta o cabeçalho de um jeito
  ("Satus" vs "Status", "Convênio/Plano" vs "Convenio/Plano", "Observação" vs "Observacao").
  A união mapeia tudo por nome normalizado, escolhe a aba certa de cada arquivo
  (ex.: a aba "RESULT" da Intermédica, que é a versão final) e ainda converte as unidades
  ("01 · Centro" → "UNIDADE 01").
- **Robôs resilientes**: tratam queda de sessão, pedem o login de novo quando precisa,
  tiram print da tela quando um elemento não aparece, e param sozinhos quando o mesmo erro
  se repete em várias guias (pra não estragar o lote inteiro).

---

## Tecnologias

- **Python 3.10+**
- **customtkinter** — interface desktop
- **Playwright** — automação dos portais num navegador real (Chromium/Chrome)
- **openpyxl** — leitura/escrita de planilhas
- **reportlab / pypdf / pdfplumber** — geração e leitura de PDFs (guias)
- **PyInstaller** — empacota tudo num `.exe` pra rodar sem instalar nada

---

## Como rodar

```bash
# 1. instalar as dependências (uma vez)
pip install -r requirements.txt
python -m playwright install chromium

# 2. abrir a central
python app.py
```

No Windows, dá pra usar os atalhos:
- `Instalar dependencias.bat` (uma vez)
- `Abrir Central de Automacoes.bat`
- `build_exe.bat` gera o `.exe` em `dist/`.

### Testar a "União" sem dados reais

A pasta [`sample_data/`](sample_data) tem planilhas **fictícias** de exemplo. Abra a
**União de conferências**, adicione os arquivos `Conferência ... dia 0X.04.xlsx` e clique em
"Juntar numa geral" — o resultado deve sair igual ao `Conferência geral (exemplo).xlsx`.

---

## Estrutura

```
app.py                     # tela inicial (um botão por automação)
interface_*.py             # as janelas de cada automação
robo_*.py                  # a lógica de cada robô (Playwright)
uniao.py                   # junta as conferências numa geral
conferencia_faturamento.py # confere faturamento x planilha x portal
feriado_engine.py          # motor da redistribuição de feriado
gerador_agenda.py          # gera as agendas remanejadas
xlsx_seguro.py             # leitura tolerante de .xlsx
navegador.py               # acha/instala o Chromium do Playwright
sample_data/               # planilhas fictícias pra testar
```

---

## Observações

- Projeto feito pra uso interno de uma clínica; as URLs dos portais são públicas, mas o fluxo
  de cada robô é específico de cada convênio.
- Os robôs **não** burlam proteções de login nem anti-bot: eles dirigem um navegador real com a
  sua própria sessão, exatamente como se você estivesse clicando.
- Nenhum dado real de paciente está neste repositório — só os exemplos fictícios de `sample_data/`.

## Licença

[MIT](LICENSE) — sinta-se à vontade pra usar como referência.
