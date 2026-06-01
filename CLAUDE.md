# Sistema de Boletos — Guia para Claude

## O que é este sistema
Sistema Flask que roda **tanto localmente (localhost:5000) quanto no Railway** (cloud) que:
1. Lê uma planilha Excel de contratos e boletos de condomínio em PDF
2. Cruza os dados via Claude AI (`claude-haiku-4-5-20251001`) e gera resumo de cobranças por locatário
3. Tem uma página separada (`/reajustes`) que calcula reajuste de aluguel usando índices do BACEN
4. Tem uma página separada (`/dimob`) que gera informes anuais de pagamento por locatário (DIMOB)
5. Tem seção de **Informes Mensais ao Proprietário** (PDF + envio por email via Resend)

## Executar localmente
```
cd sistema_boletos
python app.py
# acesse http://localhost:5000
```
Ou clicar em **`iniciar.bat`** (abre janela de terminal) /  **`iniciar_silencioso.vbs`** (sem janela — inicia automaticamente com o Windows, aguarda 10 s para OneDrive carregar).

Credenciais padrão locais:
- Admin: usuário `vitho` / senha `vi28041305`
- Usuário: usuário `lourdes` / senha `vi280407`

## Deploy Railway
```
git add .
git commit -m "descrição"
git push
```
Railway detecta o push no GitHub (`vithorafailde/sistema-boletos`) e faz deploy automático. **Não usar `railway up`** — o deploy é via GitHub.

---

## Perfis de usuário

### Admin (`vitho`)
Acesso total — vê e configura API Anthropic, email SMTP/Resend, e todas as funcionalidades.

### Usuário (`lourdes`)
Acesso total **exceto** configurações de API e Email (badges e botões ocultos no header via `{% if is_admin %}`). Rotas `/config` e `/configurar_smtp` protegidas pelo decorator `@admin_required` → retorna 403 se não-admin tentar acessar.

**Variáveis de ambiente:**
- `LOGIN_USUARIO` / `LOGIN_SENHA` → admin (padrão: vitho/vi28041305)
- `LOGIN_USUARIO_2` / `LOGIN_SENHA_2` → usuário (padrão: lourdes/vi280407)

A sessão guarda `session["role"]` = `"admin"` ou `"user"`. A função `is_admin()` verifica `session.get("role") == "admin"`. O template recebe `is_admin=is_admin()` via `render_template`.

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
| N | 13 | Email do proprietário (`email_proprietario`) — usado para envio dos informes |

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

## Informes Mensais ao Proprietário

### Como funciona
- Botão **"📄 Informes ao Proprietário"** na tela de resultados abre a seção de informes
- Select dropdown filtra por proprietário (ordem alfabética)
- Card mostra todas as unidades do proprietário com discriminação completa do repasse
- Botão **"📄 Gerar PDF"** → abre janela com PDF profissional (Times New Roman, P&B, logo Funchal)
- Botão **"📧 Enviar por E-mail"** → envia via Resend para o email da coluna N do Excel

### Cálculo do repasse (`calcRepasseData(row)`)
```
repasse = aluguel
        - taxaImob (% sobre aluguel, zero se paAtivo)
        - dedPropTotal (itens cond. marcados como encargo do prop.)
        - dedsManual (ded1..ded10, podem ser + ou -)
        - paVal (primeiro aluguel, se ativo)
        + multaBruta (multa por atraso — bruto)
        + jurosBruto (juros de mora — bruto)
        - taxaSobRec (% imob. sobre multa+juros, zero se paAtivo)
```

### Multa por atraso e juros de mora
- Campos editáveis na coluna **Repasse ao Proprietário** de cada unidade
- A taxa da imobiliária (% do contrato) é aplicada sobre multa+juros: `taxaSobRec = (multa + juros) * percImob / 100`
- Valores brutos somam ao repasse; a taxa da imob. é deduzida separadamente
- Aparecem discriminados no card de informe, no PDF e no email
- **Integração DIMOB:** ao clicar "Salvar para próximo mês", multa+juros são gravados em `dimob_historico.json` por locatário/mês. Na DIMOB, o valor do mês correspondente já inclui multa+juros automaticamente.

### Logo nos PDFs
- A logo (`static/logo.png`) é convertida para **base64 via fetch** antes de abrir a janela do PDF
- Isso garante o carregamento mesmo no Railway (evita problema de timing com `window.print()`)
- Mesmo approach na DIMOB (`gerarInformeHTML()`)

### Envio por email (Resend)
- Serviço: **Resend** (resend.com) — API HTTP, não SMTP (Railway bloqueia SMTP)
- Remetente: `Funchal Imoveis <noreply@funchalimoveis.com.br>` (domínio verificado via DNS)
- Reply-To: email configurado em "Configurar Email" → campo Usuário
- Chave API: guardada em `config.json` como `resend_api_key`
- Rota `/enviar_informe` (POST): recebe `{proprietario, email_dest, mes, rows, total}`
- Rota `/configurar_smtp` (POST, admin): salva config e testa conexão
- Rota `/get_smtp` (GET): retorna config atual + flag `configurado`
- **Railway bloqueia SMTP (465/587)** — usar sempre Resend via HTTPS

### Variáveis SMTP/Resend em config.json
```json
{
  "smtp_host": "smtp.gmail.com",
  "smtp_port": 587,
  "smtp_user": "vithor.a.failde@gmail.com",
  "smtp_pass": "",
  "smtp_from": "",
  "resend_api_key": "re_..."
}
```

---

## Banner de cota condominial divergente

- Se o valor da cota condominial no boleto do mês for diferente do histórico, exibe banner amarelo
- Campos: `row.cond_alerta`, `row.cond_cota_atual`, `row.cond_cota_hist`
- Banner tem botão fechar (X) e lista todas as unidades afetadas

---

## Bugs já corrigidos — não regredir

### 10. percImob indefinido em atualizarRodape
**Problema:** A variável `percImob` era usada em `atualizarRodape()` para calcular taxa sobre multa/juros, mas não estava definida nesse escopo (estava definida apenas em `renderRepasse()`).  
**Fix:** Adicionado `const percImobR = parseFloat(row.percentual_imob)||0` dentro do loop de `atualizarRodape()`.

### 11. MESES_PT com 13 elementos
**Problema:** A lista `MESES_PT` em `salvar_extras` tinha "março" E "marco", totalizando 13 elementos. Isso fazia todos os meses a partir de abril serem gravados com número errado (+1) no `dimob_historico.json`.  
**Fix:** Substituída por dict `{"jan":1,"fev":2,...,"dez":12}` usando os 3 primeiros caracteres do mês.

### 5. IPCA divergindo do BACEN Cidadão
**Problema:** `calcular_por_numero_indice` usava `Index(aniv_atual) / Index(aniv_ano_anterior) − 1` = 12 variações mensais, enquanto o BACEN Cidadão inclui o próprio mês de aniversário = 13 variações.  
**Fix:** Denominador alterado para o mês ANTERIOR ao aniversário do ano anterior — `Index(abr/2026) / Index(mar/2025) − 1` — igual ao IGPM (produto 13m). **Não reverter.**

### 7. "outros" é encargo do proprietário
**Decisão:** `outros` foi movido de `REPASSE_ITENS` para `NAO_REPASSE`. Qualquer item classificado como "outros" no boleto de condomínio é absorvido pelo proprietário, não repassado ao locatário. **Não reverter.**

### 8. Demonstrativo com Recibo do Pagador
**8b. Match errado (proprietário vs locatário):** PDFs no formato `condominio altavilla - joao - denis apto 37` casavam com o João (proprietário) em vez do Denis (locatário).  
**Fix:** `extrair_nome_arquivo()` reescrita — extrai o último segmento antes de "apto NNN" como nome principal (= locatário).

**8c–8e. REVERTIDO:** As tentativas de ler e lançar valores de boletos com demonstrativo foram completamente revertidas. O sistema rejeita qualquer PDF sem cv/vm no nome do arquivo (`sem_consumos`). Não tentar de novo.

### 9. "Primeiro Aluguel" no repasse ao proprietário
**Feature:** Checkbox "Primeiro Aluguel" — quando ativo, taxa de administração vai a zero e o valor informado é deduzido do repasse. Não salvo no histórico.

### 6. Filtros de endereço na página de boletos
**Decisão:** Filtros **Itaúna**, **Mandaqui**, **Ester Margareta**, **Henrique Felipe**, **Outros endereços**. Não reintroduzir filtros "Com Cond." / "Sem Cond.".

### 1–4. Bugs de reajuste
Ver versões anteriores do CLAUDE.md.

---

## Paleta de cores — decisões visuais fixas

Aplicada em `templates/index.html` e `templates/reajustes.html`. **Não alterar sem pedido explícito.**

```css
--bg:      #D6E3E8
--card:    #ffffff
--border:  #B8CECE
--text:    #1A2C30
--muted:   #4A6668
--accent:  #3B7A8E
--green:   #4A6E9A
--red:     #B84545
--yellow:  #8C6C1F
--orange:  #B86030
--header:  #2A4A52
--purple:  #6A4F8A
```

### Coluna "Repasse ao Proprietário"
- Cabeçalho (`<th>`): `background: #9B4343`
- Box de cada linha: `background: #F2DADA; border-color: #D4A0A0`

---

## O que NÃO adicionar sem pedido explícito

- ❌ Ler/processar boletos com demonstrativo (sem cv/vm no nome do arquivo)
- ❌ Mover `outros` de volta para `REPASSE_ITENS`
- ❌ Filtros "Com Cond." / "Sem Cond." na página de boletos
- ❌ Mensagem "Planilha encontrada" na página `/reajustes`
- ❌ Botão "Exportar CSV" na página de reajustes
- ❌ Verificação de "vigência mínima de 12 meses"
- ❌ Fonte alternativa de dados (IPEA, FGV) para qualquer índice
- ❌ Janela de 12 meses (decisão: 13 meses — BACEN Cidadão)
- ❌ Coluna "VENCIDO X dias" ou qualquer conceito de atraso
- ❌ Alterar a coluna H (data de aniversário) ao aplicar reajuste
- ❌ SMTP direto no Railway (portas 465/587 bloqueadas) — usar sempre Resend via HTTPS

---

## Arquivos principais

```
app.py                        — servidor Flask, toda a lógica backend
templates/home.html           — landing page com cards: Boletos / Reajustes / DIMOB
templates/index.html          — página de boletos + informes mensais
templates/reajustes.html      — página de reajuste de aluguel
templates/dimob.html          — página DIMOB (informes anuais)
templates/login.html          — login
static/logo.png               — logo Funchal Imóveis (usada nos PDFs em base64)
.railwayignore                — exclui uploads/ e data/config.json do deploy
Procfile                      — gunicorn, 1 worker, timeout 300s
iniciar.bat                   — inicia o servidor local (janela visível)
iniciar_silencioso.vbs        — inicia o servidor local sem janela (auto-start Windows)
data/config.json              — API key + config SMTP/Resend (criado pela UI)
data/historico.json           — histórico de IPTU, seguros e extras por locatário
data/dimob_historico.json     — histórico de reajustes + multa/juros para o DIMOB
```

## Variáveis de ambiente no Railway
- `SECRET_KEY` — chave Flask
- `LOGIN_USUARIO` / `LOGIN_SENHA` — credenciais admin
- `LOGIN_USUARIO_2` / `LOGIN_SENHA_2` — credenciais usuário (Lourdes)
- `ANTHROPIC_API_KEY` — chave da API Claude (sobrepõe data/config.json)
- `SMTP_HOST/PORT/USER/PASS/FROM` — config SMTP (sobrepõe config.json)
- `RESEND_API_KEY` — chave Resend (sobrepõe config.json)

## Rotas Flask completas
| Rota | Método | Acesso | Descrição |
|------|--------|--------|-----------|
| `/login` | GET/POST | público | Autenticação |
| `/logout` | GET | logado | Encerra sessão |
| `/` | GET | logado | Landing page |
| `/boletos` | GET | logado | Página de boletos |
| `/config` | POST | **admin** | Salva API key |
| `/processar` | POST | logado | Processa planilha + PDFs |
| `/salvar_extras` | POST | logado | Salva histórico + multa/juros no DIMOB |
| `/baixar_historico` | GET | logado | Download historico.json |
| `/restaurar_historico` | POST | logado | Restaura historico.json |
| `/exportar` | POST | logado | Exporta Excel |
| `/configurar_smtp` | POST | **admin** | Salva e testa config email |
| `/get_smtp` | GET | logado | Retorna config email atual |
| `/enviar_informe` | POST | logado | Envia informe por email (Resend) |
| `/reajustes` | GET | logado | Página de reajuste |
| `/api/calcular_reajustes` | POST | logado | Calcula reajustes via BACEN |
| `/api/aplicar_reajustes` | POST | logado | Grava novos aluguéis na planilha |
| `/api/baixar_contratos` | GET | logado | Download planilha atualizada |
| `/dimob` | GET | logado | Página DIMOB |
| `/api/dimob_calcular` | POST | logado | Calcula informes anuais |
| `/api/dimob_salvar_historico` | POST | logado | Salva aluguéis anteriores |
| `/api/dimob_exportar` | POST | logado | Exporta Excel DIMOB |

---

## Módulo DIMOB — Regras que NÃO podem mudar

### Lógica de divisão de meses (regra central)
```
mes_aplicacao = data_rej.month % 12 + 1   ← mês SEGUINTE ao aniversário
```
- Meses **antes** de `mes_aplicacao` → `aluguel_antigo`
- Mês `mes_aplicacao` **em diante** → `aluguel_atual`
- Contrato inativo → `None`

### Multa/juros no DIMOB
- Gravados em `dimob_historico.json` quando usuário clica "Salvar para próximo mês"
- Estrutura: `hist[ano][chave_locatario]["multa_juros"][str(mes_num)] = {"multa": X, "juros": Y}`
- `calcular_meses_dimob()` soma multa+juros ao valor base de cada mês automaticamente
- Mês extraído do campo `mes` do payload (ex: "Junho/2026" → 6) via dict `{"jan":1,...,"dez":12}`

### Estrutura do dimob_historico.json
```json
{
  "2026": {
    "LOCATARIONORM": {
      "aluguel_antigo": 1800.00,
      "mes_aplicacao": 2,
      "multa_juros": {
        "6": {"multa": 150.0, "juros": 15.0}
      }
    }
  }
}
```

### PDF por unidade
Gerado no cliente (JavaScript): `window.open()` + `window.print()`. Logo carregada via fetch→base64.

---

## Constantes importantes em app.py

```python
REPASSE_ITENS   = ["agua", "gas", "energia", "tx_leitura", "vaga_moto", "tag"]
NAO_REPASSE     = ["fundo_reserva", "correio", "melhorias", "outros"]
EDIFICIOS_EXCLUIDOS = ["lamelas", "paisagem", "victoria"]
```

## localStorage (persistência no browser)
- `boletos_dados` — JSON dos locatários processados
- `boletos_mes` — mês de referência
- `boletos_salvo_em` — "AAAA-MM" do momento do processamento
- Restore automático ao carregar a página; aceita dados do mês atual ou anterior (diff ≤ 1 mês)
