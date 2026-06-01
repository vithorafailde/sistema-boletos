# Sistema de Boletos — Guia para Claude

## O que é este sistema
Sistema Flask que roda **tanto localmente (localhost:5000) quanto no Railway** (cloud) que:
1. Lê uma planilha Excel de contratos e boletos de condomínio em PDF
2. Cruza os dados via Claude AI (`claude-haiku-4-5-20251001`) e gera resumo de cobranças por locatário
3. Tem uma página separada (`/reajustes`) que calcula reajuste de aluguel usando índices do BACEN
4. Tem uma página separada (`/dimob`) que gera informes anuais de pagamento por locatário (DIMOB)

## Executar localmente
```
cd sistema_boletos
python app.py
# acesse http://localhost:5000
```
Ou clicar em **`iniciar.bat`** (abre janela de terminal) /  **`iniciar_silencioso.vbs`** (sem janela — inicia automaticamente com o Windows, aguarda 10 s para OneDrive carregar).

Credenciais padrão locais: usuário `vitho` / senha `vi28041305` (substituíveis pelas env vars `LOGIN_USUARIO` / `LOGIN_SENHA`).

## Deploy Railway
```
railway up --detach
```
Sempre usar `--detach`.

---

## Estrutura da planilha `contratos.xlsx`

| Coluna | Índice (0-based) | Conteúdo |
|--------|-----------------|----------|
| A | 0 | tipo |
| B | 1 | tipo2 |
| C | 2 | endereço |
| D | 3 | proprietário |
| E | 4 | locatário |
| F | 5 | aluguel (float) |
| G | 6 | percentual_imob (comissão %) |
| H | 7 | data_reajuste (formato DD/MM, sem ano) |
| I | 8 | data_inicio |
| J | 9 | data_fim |
| K | 10 | índice (IPCA / IGPM / INPC) |

| L | 11 | CPF do proprietário (pode ser float no Excel — usar `_cpf()` para limpar) |
| M | 12 | CPF do locatário (mesmo tratamento) |

A mesma planilha é usada pelos três sistemas (boletos, reajustes e DIMOB).  
`ler_excel()` → sistema de boletos  
`ler_excel_reajustes()` → sistema de reajustes  
`ler_excel_dimob()` → sistema DIMOB

---

## Sistema de Reajustes — Regras que NÃO podem mudar

### Janela de cálculo: 12 variações mensais (mês seguinte ao aniversário anterior até o aniversário atual)
- A janela é de **12 variações mensais**: do mês M+1 do ano anterior até o mês M do ano atual
- Ex.: aniversário em abril/2026 → janela **mai/2025 a abr/2026** (12 variações)
- A variação de abr/2025 representa mar→abr/2025 e **não pertence** ao período abr/2025→abr/2026
- Equivale ao cálculo do BACEN Cidadão
- Função: `calcular_acumulado_12m()` usa `range(11, -1, -1)` — **NÃO alterar para 12 ou 13**

### Fonte dos índices — duas estratégias por status

**Fonte dos índices por série BACEN SGS:**
- IGPM: **série 28655** (variação mensal com ~10 casas decimais) — dá o mesmo resultado que o BACEN Cidadão
  - NÃO usar série 189 para IGPM — publica com apenas 2 decimais, acumula erro de ~0.02% em 13 meses
- IPCA mensal: série 433 | INPC mensal: série 188

**Contratos ESTE_MES:**
- IGPM: produto 13m com série 28655 → `calcular_acumulado_12m(historicos_mensal['IGPM'], data_rej)`
- IPCA: número-índice BACEN série **1737** via `calcular_por_numero_indice`
- INPC: número-índice BACEN série **1617** via `calcular_por_numero_indice`
- `calcular_por_numero_indice` usa denominador = mês ANTERIOR ao aniversário do ano anterior
  - Ex.: aniversário abr/2026 → Index(2026-04) / Index(2025-03) − 1 (13 variações, igual BACEN Cidadão)
  - **NÃO usar** Index(aniv_ano_anterior) como denominador — isso dá só 12 variações e diverge do BACEN
- Retorna `meses_base = 12` quando dados completos

**Contratos FUTURO** → variação mensal composta (dados parciais):
- IGPM = série 28655, IPCA = série 433, INPC = série 188
- Fórmula: produto dos % mensais disponíveis
- Retorna `meses_base = X` (X < 12, mostra aviso ⚠)

**Fallback**: se número-índice falhar para ESTE_MES, usa % mensal composto

### Status dos contratos
- `ESTE_MES` → mês atual == mês SEGUINTE ao aniversário (hora de reajustar)
- `FUTURO` → mes_aplicacao > hoje.month
- `OK` → mes_aplicacao ≤ hoje.month (já foi reajustado este ciclo)
- Contratos OK: **não calcular** acumulado, não mostrar novo aluguel, não mostrar diferença
- Lógica correta: `mes_aplicacao = data_rej.month % 12 + 1`

### Aplicação do reajuste
- Só grava **coluna F** (aluguel) na planilha — nenhuma outra coluna é alterada
- A data de aniversário (col H) **não muda** — é sempre DD/MM repetida todo ano
- Escrita atômica: salva em `.tmp` depois renomeia

### Busca BACEN
- Timeout: **30 segundos**
- **3 tentativas** com 2s de intervalo entre elas
- Retorna `({}, erro_str)` se todas as tentativas falharem

---

## Bugs já corrigidos — não regredir

### 5. IPCA divergindo do BACEN Cidadão
**Problema:** `calcular_por_numero_indice` usava `Index(aniv_atual) / Index(aniv_ano_anterior) − 1` = 12 variações mensais, enquanto o BACEN Cidadão inclui o próprio mês de aniversário = 13 variações.  
**Fix:** Denominador alterado para o mês ANTERIOR ao aniversário do ano anterior — `Index(abr/2026) / Index(mar/2025) − 1` — igual ao IGPM (produto 13m). **Não reverter.**

### 7. "outros" é encargo do proprietário
**Decisão:** `outros` foi movido de `REPASSE_ITENS` para `NAO_REPASSE`. Qualquer item classificado como "outros" no boleto de condomínio é absorvido pelo proprietário, não repassado ao locatário. **Não reverter.**

### 8. Demonstrativo com Recibo do Pagador
**Problema:** Boletos que vêm com balancete demonstrativo do prédio inteiro + seção "Composição da Arrecadação" / "Discriminação das Verbas" com valores individuais da unidade eram ignorados pelo sistema. Três sub-problemas corrigidos:

**8a. Classificação errada:** Claude classificava como `demonstrativo_geral` mesmo havendo seção individual.  
**Fix:** Bloco `ATENCAO` adicionado ao `PROMPT_CONDO`: se o texto tiver dados do prédio inteiro E uma seção "Discriminação das Verbas" / "Recibo do Pagador" com valores individuais, classificar como `demonstrativo_com_recibo`. Nunca classificar como `demonstrativo_geral` só por ter demonstrativo geral junto.

**8b. Match errado (proprietário vs locatário):** PDFs no formato `condominio altavilla - joao - denis apto 37` casavam com o João (proprietário) em vez do Denis (locatário).  
**Fix:** `extrair_nome_arquivo()` reescrita — extrai o último segmento antes de "apto NNN" como nome principal (= locatário). O step `ac_palavras` usa `score_nome_arquivo` com s\*98 para aceitar nomes parciais do campo A/C. O step `nome_arquivo_prop` mantém s\*87 para não sobrepor o match de locatário.

**8c–8e. REVERTIDO:** As tentativas de ler e lançar valores de boletos com demonstrativo (Denis/Altavilla e similares) foram completamente revertidas. O sistema volta a rejeitar qualquer PDF sem cv/vm no nome do arquivo (`sem_consumos`). Demonstrativos continuam sendo ignorados como antes. A única alteração mantida é a 8b (fix do match locatário/proprietário no nome do arquivo).

### 9. "Primeiro Aluguel" no repasse ao proprietário
**Feature:** Checkbox "Primeiro Aluguel" na coluna de Repasse ao Proprietário de cada unidade. Quando marcado e preenchido com valor: a taxa de administração vai a zero e o valor informado é deduzido do repasse. Não é salvo no histórico (evento único). Também corrigido bug em `atualizarRodape()` que não somava ded5 e ded6.

### 6. Filtros de endereço na página de boletos
**Decisão:** Os botões "Com Cond." e "Sem Cond." foram **removidos** e substituídos por filtros de endereço: **Itaúna**, **Mandaqui**, **Outros endereços**. Não reintroduzir os filtros de condomínio.  
A classificação usa `row.endereco.toLowerCase()` com `includes('itauna'/'itaúna'/'mandaqui')`. Todo endereço que não for nenhum dos dois vai para "Outros endereços".

### 1. File overwrite após aplicar
**Problema:** Após aplicar reajustes, o código chamava `calcular()` que re-enviava o arquivo original (ainda em `arquivoSelecionado`), sobrescrevendo a planilha atualizada.  
**Fix:** Após aplicar, zerar `arquivoSelecionado = null` e **não** chamar `calcular()`. Atualizar estado local do `todosContratos` diretamente.

### 2. painelDownload sumindo
**Problema:** `painelDownload` estava dentro de `painelAplicar`. Quando `calcular()` ocultava `painelAplicar` (porque `aplicaveis == 0`), o botão de download sumia junto.  
**Fix:** `painelDownload` fica **fora** de `painelAplicar` no HTML.

### 3. IPCA não calculando no Railway
**Problema:** Timeout de 15s era insuficiente no Railway, IPCA falhava silenciosamente.  
**Fix:** Timeout de 30s + 3 tentativas com retry.

### 4. Double-apply
**Problema:** Após aplicar, recalcular mostrava o novo aluguel como base para outro reajuste.  
**Fix:** Contratos aplicados são marcados como `status: 'OK'` localmente, sem nova chamada ao BACEN.

---

## Paleta de cores — decisões visuais fixas

Aplicada em `templates/index.html` e `templates/reajustes.html`. **Não alterar sem pedido explícito.**

```css
--bg:      #D6E3E8   /* fundo geral — cinza-teal com vida */
--card:    #ffffff
--border:  #B8CECE
--text:    #1A2C30   /* texto principal — escuro, boa legibilidade */
--muted:   #4A6668
--accent:  #3B7A8E   /* teal-azul — botões, links, foco */
--green:   #4A6E9A   /* slate-índigo — sem verde intencional */
--red:     #B84545
--yellow:  #8C6C1F
--orange:  #B86030
--header:  #2A4A52   /* header das páginas */
--purple:  #6A4F8A
```

### Coluna "Repasse ao Proprietário"
- Cabeçalho (`<th>`): `background: #9B4343` (laranja-avermelhado saturado)
- Box de cada linha: `background: #F2DADA; border-color: #D4A0A0` (tom mais fraco do mesmo vermelho)

### Stats (página de boletos)
- Todos os números usam a cor padrão `--text` — **sem cores individuais por stat**
- Exceção: "PDFs sem match" usa `--red` quando há ocorrências (é um alerta)
- A caixa "Comissão Imob." usa o mesmo estilo dos outros stats (sem borda/fundo especial)

### Botões de filtro
- Todos usam `btn-outline` sem `style` inline — **sem cores individuais por botão**
- Ativo: fundo `--accent`

---

## O que NÃO adicionar sem pedido explícito

- ❌ Ler/processar boletos com demonstrativo (sem cv/vm no nome do arquivo) — foi tentado e revertido. Claude não consegue extrair valores individuais com confiança dessas páginas. Esses boletos são rejeitados com `rejeitado: "sem_consumos"`. Não tentar de novo.
- ❌ Mover `outros` de volta para `REPASSE_ITENS` — é encargo do proprietário. Não reverter.
- ❌ Filtros "Com Cond." / "Sem Cond." na página de boletos — foram substituídos pelos filtros de endereço (Itaúna / Mandaqui / Outros endereços). Não restaurar.
- ❌ Mensagem "Planilha encontrada" na página `/reajustes` — foi removida intencionalmente. Não reintroduzir via Jinja2 (`tem_excel`), JavaScript, ou qualquer outro meio. A rota `reajustes_page()` passa `tem_excel=False` fixo e o template não deve ter bloco `{% if tem_excel %}`.
- ❌ Botão "Exportar CSV" na página de reajustes (foi removido intencionalmente — confundia o usuário)
- ❌ Verificação de "vigência mínima de 12 meses" (não foi pedida e quebra contratos válidos)
- ❌ Fonte alternativa de dados (IPEA, FGV) para qualquer índice
- ❌ Janela de 12 meses (decisão: 13 meses — BACEN Cidadão)
- ❌ Coluna "VENCIDO X dias" ou qualquer conceito de atraso — só ESTE_MES / FUTURO / OK
- ❌ Alterar a coluna H (data de aniversário) ao aplicar reajuste

---

## Arquivos principais

```
app.py                        — servidor Flask, toda a lógica backend
templates/home.html           — landing page com cards: Boletos / Reajustes / DIMOB
templates/index.html          — página de boletos
templates/reajustes.html      — página de reajuste de aluguel
templates/dimob.html          — página DIMOB (informes anuais)
templates/login.html          — login
.railwayignore                — exclui uploads/ e data/config.json do deploy
Procfile                      — gunicorn, 1 worker, timeout 300s
iniciar.bat                   — inicia o servidor local (janela visível)
iniciar_silencioso.vbs        — inicia o servidor local sem janela (auto-start Windows)
gerar_relatorio.py            — script standalone (não Flask) para gerar relatório HTML de conferência
test_dimob.py                 — testes standalone do módulo DIMOB (rodar fora do Flask)
data/config.json              — armazena a API key localmente (criado pela UI em /config)
data/historico.json           — histórico de IPTU, seguros e extras por locatário
data/dimob_historico.json     — histórico de reajustes para o DIMOB (criado automaticamente)
```

## Variáveis de ambiente no Railway
- `SECRET_KEY` — chave Flask
- `LOGIN_USUARIO` / `LOGIN_SENHA` — credenciais de acesso
- `ANTHROPIC_API_KEY` — chave da API Claude (sobrepõe data/config.json)

## API key — duas fontes
- **Local**: salva em `data/config.json` via rota `/config` da UI
- **Railway**: `ANTHROPIC_API_KEY` env var (tem prioridade sobre config.json)
- `ler_config()` retorna a env var se existir; caso contrário usa config.json

## UPLOAD_DIR
- Fica em `uploads/` (relativo ao app.py)
- A rota `/processar` apaga todos os arquivos de `uploads/` **exceto `.xlsx`** — preserva o `contratos.xlsx` para a página de reajustes
- `contratos.xlsx` é o nome fixo da planilha dentro de `uploads/`
- `_excel_path_contratos()` busca `contratos.xlsx` primeiro, depois qualquer `.xlsx`

---

## aviso12 na tabela de reajustes
- `12/12m` em cinza → todos os 12 meses encontrados (dados completos)
- `⚠ X/12m` em laranja → dados parciais (meses futuros sem dado no BACEN ainda)
- nada → contrato OK (não calculado)

---

## Constantes importantes em app.py

```python
REPASSE_ITENS   = ["agua", "gas", "energia", "tx_leitura", "vaga_moto", "tag"]
NAO_REPASSE     = ["fundo_reserva", "correio", "melhorias", "outros"]   # encargos do proprietário
EDIFICIOS_EXCLUIDOS = ["lamelas", "paisagem", "victoria"]     # só aluguel, sem repasse cond
EDIFICIOS_CONSUMO_KEYWORDS = ["casa verde", "mandaqui", ...]  # têm água/gás medidos por unidade
```

- `EDIFICIOS_EXCLUIDOS`: imóveis onde o locatário paga **só aluguel** — nenhum custo de condomínio é repassado
- `REPASSE_ITENS` / `NAO_REPASSE`: controlam quais itens do boleto de condomínio aparecem na cobrança do locatário

## vigencia_ok
- Calculado em `ler_excel_reajustes()`: `(aniv - data_inicio).days >= 365`
- Se `data_inicio` ausente, assume vigência OK
- O campo existe no dict do contrato, mas a UI pode exibir aviso — **não bloqueia** o reajuste automaticamente, apenas informa

## Rotas Flask completas
| Rota | Método | Descrição |
|------|--------|-----------|
| `/login` | GET/POST | Autenticação |
| `/logout` | GET | Encerra sessão |
| `/` | GET | Landing page (cards: Boletos / Reajustes / DIMOB) |
| `/boletos` | GET | Página de boletos (era `/`) |
| `/config` | POST | Salva API key |
| `/processar` | POST | Processa planilha + PDFs via Claude |
| `/salvar_extras` | POST | Salva IPTU/seguros/extras no histórico |
| `/baixar_historico` | GET | Download de historico.json |
| `/restaurar_historico` | POST | Restaura historico.json |
| `/exportar` | POST | Exporta resultado como .xlsx |
| `/reajustes` | GET | Página de reajuste de aluguel |
| `/api/calcular_reajustes` | POST | Calcula reajustes via BACEN/IPEA |
| `/api/aplicar_reajustes` | POST | Grava novos aluguéis na planilha |
| `/api/debug_igpm` | GET | Debug: busca bruta do IGPM no IPEA |
| `/api/baixar_contratos` | GET | Download da planilha atualizada |
| `/dimob` | GET | Página DIMOB |
| `/api/dimob_calcular` | POST | Calcula informes anuais por unidade |
| `/api/dimob_salvar_historico` | POST | Salva aluguéis anteriores manualmente |
| `/api/dimob_exportar` | POST | Exporta Excel com um sheet por proprietário |

---

## Módulo DIMOB — Regras que NÃO podem mudar

### O que é
Declaração de Informações sobre Atividades Imobiliárias. Para cada contrato, exibe o valor pago mês a mês durante o ano, dividindo corretamente o período antes e depois do reajuste.

### Lógica de divisão de meses (regra central)
```
mes_aplicacao = data_rej.month % 12 + 1   ← mês SEGUINTE ao aniversário
```
- Meses **antes** de `mes_aplicacao` → `aluguel_antigo` (valor antes do reajuste)
- Mês `mes_aplicacao` **em diante** → `aluguel_atual` (valor atual na coluna F)
- Contrato inativo em determinado mês (antes de `data_inicio` ou após `data_fim`) → `None` (não contabilizado)
- **NÃO alterar esta lógica** — é a mesma fórmula usada em `/reajustes`

Exemplo: aniversário setembro (col H = `01/09`), aluguel antigo R$ 1.800, atual R$ 2.000:

| Jan–Set | Out–Dez |
|---------|---------|
| 1.800   | 2.000   |

### Fonte do aluguel_antigo — duas origens
1. **Automática:** quando reajuste é aplicado em `/reajustes`, `aplicar_reajustes_excel()` salva o valor anterior em `data/dimob_historico.json` ANTES de sobrescrever a coluna F.
2. **Manual:** usuário digita na coluna "Alug. Anterior" na tabela DIMOB para contratos sem histórico. O sistema usa o `mes_aplicacao` da planilha (col H) para saber onde dividir.

**Requisito:** para que a divisão manual funcione, o contrato PRECISA ter a data de reajuste preenchida na coluna H. Sem ela, `mes_aplicacao = None` e o sistema exibe todos os meses com o aluguel atual (sem divisão).

### Estrutura do dimob_historico.json
```json
{
  "2025": {
    "nome_locatario_normalizado": {
      "aluguel_antigo": 1800.00,
      "aluguel_novo":   2000.00,
      "locatario":      "NOME ORIGINAL",
      "mes_aplicacao":  10,
      "num_linha":      5
    }
  }
}
```
- Chave: `norm(locatario)` — minúsculas, sem acentos, sem espaços duplos
- **Não está no `.railwayignore`** — é enviado ao Railway no deploy (diferente de `historico.json` que está excluído)

### PDF por unidade
Gerado inteiramente no cliente (JavaScript): `window.open()` + `window.print()`. Sem dependências Python novas (sem reportlab/weasyprint). A função `gerarInformeHTML()` usa `escHtml()` para escapar todos os campos de dados do usuário.

### Excel exportado
- Um sheet por proprietário (nome truncado a 31 chars, sufixo `_1`/`_2` se duplicado)
- Formato "INFORME DE PAGAMENTOS" com cabeçalho, grid de meses, total anual e comissão
- `PatternFill("solid", start_color="FFFFFF")` — necessário para compatibilidade com todas as versões do openpyxl

### CPF como float no Excel
openpyxl lê CPF numérico como `12345678900.0`. Usar sempre a função auxiliar `_cpf(v)` definida dentro de `ler_excel_dimob()` para converter corretamente (remove `.0` final se o restante for só dígitos).
