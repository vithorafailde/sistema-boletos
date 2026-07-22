# Sistema de Boletos — Guia para Claude

## O que é este sistema
Sistema Flask que roda **tanto localmente (localhost:5000) quanto no Railway** (cloud) que:
1. Lê uma planilha Excel de contratos e boletos de condomínio em PDF
2. Cruza os dados via Claude AI (`claude-sonnet-4-6`) e gera resumo de cobranças por locatário
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
- `RENOVAR` → `data_fim` (col J) caiu no mês passado — contrato venceu, precisa renovação
- `ESTE_MES` → mês atual == mês SEGUINTE ao aniversário (hora de reajustar)
- `FUTURO` → mes_aplicacao > hoje.month
- `OK` → mes_aplicacao ≤ hoje.month (já foi reajustado este ciclo)
- Contratos OK: **não calcular** acumulado, não mostrar novo aluguel, não mostrar diferença
- Lógica correta: `mes_aplicacao = data_rej.month % 12 + 1`
- Ordenação: RENOVAR > ESTE_MES > FUTURO > OK

### Status RENOVAR — regras que NÃO podem mudar
- Detectado em `ler_excel_reajustes()` **independente** do status de reajuste
- Condição: `data_fim.month == hoje.month - 1` AND `data_fim.year == ano_anterior`
- Ex.: `data_fim = 01/05/2026`, hoje = junho/2026 → RENOVAR ✓
- Ex.: `data_fim = 01/06/2026`, hoje = junho/2026 → ainda não (contrato acabou de encerrar)
- **Campo manual de valor** ("Novo valor do contrato") aparece **apenas** para RENOVAR — não para ESTE_MES
- Badge roxo "Renovar contrato" na tabela; filtro dedicado na barra de filtros

### Aplicação do reajuste
- Só grava **coluna F** (aluguel) na planilha — nenhuma outra coluna é alterada
- A data de aniversário (col H) **não muda** — é sempre DD/MM repetida todo ano
- Escrita atômica: salva em `.tmp` depois renomeia

### Busca BACEN
- Timeout: **30 segundos**
- **3 tentativas** com 2s de intervalo entre elas
- Retorna `({}, erro_str)` se todas as tentativas falharem

---

## Leitura de PDFs de condomínio (Claude AI)

- Modelo: **`claude-sonnet-4-6`** — NÃO usar Haiku, precisão muito menor
- Resolução das imagens: **250 DPI** — NÃO reduzir abaixo disso
- Modo **híbrido**: envia texto extraído pelo pdfplumber + imagens juntos para cruzamento visual
  - Mesmo quando o PDF tem texto selecionável, as imagens são incluídas
  - Instrução ao modelo: se texto e imagem divergirem, prevalece a imagem
- `max_tokens`: 2000
- Verificação de soma: o prompt instrui o modelo a somar os itens e comparar com o total do boleto
- Formato brasileiro: prompt instrui explicitamente (vírgula = decimal, ponto = milhar)

---

## Envio de Boletos por Email (página `/envio_boletos`)

### Identificação de locatário pelo nome do arquivo
- Função: `_match_nome_arquivo(nome_arq)` em `processar_boletos_locatarios()`
- Scoring: **cobertura do nome do locatário no arquivo** (não o contrário)
  - Mede quantas palavras significativas do NOME aparecem no filename
  - Ex.: locatário "Isidro Silva" em "isidro silva av paulista 101 apto 5" → 2/2 = 100%
  - Lógica antiga media palavras do arquivo → "isidro silva" em filename de 8 palavras = 25% → falha
  - **NÃO reverter para lógica antiga**
- Threshold: **40%** (antes era 50% — reduzido para tolerar sobrenomes ausentes no filename)
- Match exato (norm completo) tem prioridade sobre match por palavras

### Envio consolidado por locatário
- Locatários com **múltiplos imóveis** recebem **um único email** com todos os PDFs em anexo
- `enviarTodos()`: agrupa por `(locatario_match, email)` antes de enviar
- `enviarUm()`: ao clicar "Enviar" em qualquer linha, consolida todos os boletos não enviados do mesmo locatário+email
- Backend `/envio_boletos/enviar`: aceita `arquivos: [{arquivo_salvo}]` (array) — retrocompatível com `arquivo_salvo` (string)
- Assunto com múltiplos: "Boletos de Aluguel — Mês — Nome (N imóveis)"

### Log de auditoria
- Status gravado: sempre `"sent"` (tanto Resend quanto SMTP) — **NÃO usar "enviado"**
- `AUDIT_BADGES` no frontend mapeia: `sent`, `delivered`, `delayed`, `bounced`, `complained`, `opened`, `clicked`, `erro`

---

## Informes Mensais ao Proprietário

### Como funciona
- Card **"Informes"** na home page (`/`) leva para `/boletos?informes=1` — se houver dados em localStorage, abre direto na seção de informes; caso contrário mostra tela de upload
- **Não há botão "Informes ao Proprietário" na tela de resultados** — acesso só pelo card da home
- Select dropdown filtra por proprietário (ordem alfabética)
- Card mostra todas as unidades do proprietário com discriminação completa do repasse
- Botão **"Gerar PDF"** → abre janela com PDF profissional (Times New Roman, P&B, logo Funchal)
- Botão **"Enviar por E-mail"** → envia via Resend para o email da coluna N do Excel
- Botão **"Enviar para Todos"** → envia sequencialmente para todos os proprietários com email cadastrado; mostra progresso e resumo final de sucesso/erros; proprietários sem email são listados e ignorados
- Botão **"Editar E-mails"** → modal para editar emails dos proprietários sem precisar reupar a planilha

### Arquivo mensal de Informes — reabrir/reenviar meses passados
- Select **"Mês de Referência"** no topo da seção de Informes, ao lado do select de proprietário
- Ao clicar em **"Salvar"** na tela de Boletos, além do que já salvava (historico.json + dimob_historico.json), agora também arquiva o processamento **completo** do mês (aluguel, condomínio, IPTU, seguros, abono, deduções, comissão) em `data/informes_historico.json`, indexado por texto do mês (ex: `"Julho/2026"`)
- Rotas: `POST /api/informes_salvar_mes` (grava/sobrescreve o mês), `GET /api/informes_meses_salvos` (lista meses arquivados, mais recente primeiro), `GET /api/informes_carregar_mes?mes=X` (retorna os dados daquele mês)
- Selecionar um mês arquivado no dropdown troca `_infGrupos`/`_infOrdem` pra usar os dados arquivados (função `montarGruposInforme`) em vez dos dados "ao vivo" da tela de Boletos — reseta a seleção de proprietário ao trocar
- `mesInformeAtual()` decide qual "mes" mandar pro backend (email/PDF/log): o mês arquivado selecionado, ou o mês ao vivo se nenhum arquivado estiver selecionado — **usar sempre essa função**, nunca ler `#inMes` direto dentro do fluxo de Informes, senão o email sai com o mês errado quando se está vendo um mês arquivado
- `data/informes_historico.json` está no `.railwayignore` (igual ao `historico.json`) — **não remover de lá**, senão cada deploy apaga os meses arquivados em produção
- Só funciona pros meses salvos **depois** dessa feature existir — meses anteriores não têm registro nesse arquivo

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

### Comissão sobre o líquido — Manoel Failde, Amandio Failde, Izabel Failde
- **Pedido explícito:** para esses 3 proprietários (só eles), a comissão da imobiliária é calculada sobre `(aluguel - abono)`, não sobre o aluguel bruto. Para todos os outros proprietários, continua `aluguel * percentual_imob / 100` (sem mudança)
- **Match por palavras, não nome exato:** a planilha pode ter nome completo com nome do meio (ex.: "Manoel do Nascimento Failde") diferente do nome curto usado na regra. Por isso o match verifica se as palavras-chave (primeiro nome + sobrenome, ex. `["manoel","failde"]`) estão TODAS presentes entre as palavras do nome do proprietário — não faz `==` de string inteira. **Não trocar para comparação de nome exato** — já quebrou uma vez por causa do nome do meio
- Lista duplicada em dois lugares que precisam ficar sincronizados:
  - `app.py`: `PROPRIETARIOS_COMISSAO_LIQUIDA_PALAVRAS` (lista de sets de `norm_palavras(nome)`) + `e_proprietario_comissao_liquida(nome)` — usado em `calcular_meses_dimob`
  - `templates/index.html`: `PROPRIETARIOS_COMISSAO_LIQUIDA` (lista de arrays de palavras) + `_nomeContemPalavras()` — usado em `renderRepasse`, `atualizarRodape`, `calcRepasseData` via `temComissaoLiquida()`
- **Boletos/Informes Mensais:** `taxaImob` é recalculado no JS como `(aluguel - abono_val) * percImob/100` para esses 3, em vez de usar o `row.taxa_imob` vindo pronto do backend (que é sempre bruto)
- **DIMOB/Informe Anual:** o abono passou a ser salvo por locatário/mês em `dimob_historico.json` (chave `"abono"`, mesmo formato de `"multa_juros"`) — gravado em `salvar_extras` sempre que houver valor, pra qualquer locatário (grátis, não usado por quem não é um dos 3). `calcular_meses_dimob()` só desconta esse abono do total do mês quando `e_proprietario_comissao_liquida(proprietario)` é verdadeiro — pra esses 3, **tanto o total anual quanto a comissão** ficam líquidos de abono; pros demais nada muda
- `dimob.html`/`informe_anual.html` não precisaram de nenhuma mudança — eles recalculam comissão em cima de `c.total`/`c.meses` que já vêm líquidos do backend
- **Não expandir a lista de nomes nem mudar a fórmula sem pedido explícito**

### Multa por atraso e juros de mora
- Campos editáveis na coluna **Repasse ao Proprietário** de cada unidade
- A taxa da imobiliária (% do contrato) é aplicada sobre multa+juros: `taxaSobRec = (multa + juros) * percImob / 100`
- Valores brutos somam ao repasse; a taxa da imob. é deduzida separadamente
- Aparecem discriminados no card de informe, no PDF e no email
- **Integração DIMOB:** ao clicar "Salvar", multa+juros são gravados em `dimob_historico.json` por locatário/mês. Na DIMOB, o valor do mês correspondente já inclui multa+juros automaticamente.
- Email de informe: `_gerar_html_email` renderiza linhas separadas de multa, juros e taxa sobre eles quando presentes — o payload do frontend deve incluir `multa_atraso`, `juros_mora`, `taxa_sob_rec`

### Deduções manuais (`ded1` a `ded10`)
- Cada dedução tem `desc`, `val` e `subtrair` (bool: true = subtrai do repasse, false = soma)
- **`subtrair` é salvo no `historico.json`** e restaurado no mês seguinte — NÃO remover do `salvar_extras`
- Sem o campo `subtrair` no histórico, deduções marcadas como "+" viram "-" no próximo mês

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
- `_enviar_via_resend()` retorna o `resend_id` da resposta da API (usado na auditoria)

### Auditoria de Envios
- Todo envio (informe ao proprietário ou boleto ao locatário) é gravado em `data/log_envios.json`
- Função: `gravar_log_envio(nome, email, mes, status, erro="", tipo="informe", resend_id="")`
- Campo `tipo`: `"informe"` (página de informes) ou `"boleto"` (página envio de boletos)
- Status iniciais: `"sent"` (Resend aceitou) ou `"erro"` (falha)
- Status atualizados via webhook: `delivered`, `delayed`, `bounced`, `complained`, `opened`, `clicked`
- Rota `/api/resend_webhook` (POST, sem auth) — recebe eventos do Resend e atualiza o log pelo `resend_id`
- Rota `/api/log_envios` (GET, logado) — aceita `?tipo=informe` ou `?tipo=boleto`
- Limite: **2000 entradas** (acumulado permanente, mês a mês)
- Tabela de auditoria aparece nas páginas de Informes e Envio de Boletos, com filtro por mês
- **Para ativar status automáticos:** configurar webhook no Resend → URL: `https://[railway-url]/api/resend_webhook`

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

## Toast "Salvo!"

- Elemento `<div id="toastSalvo">` fixo no canto inferior direito, presente em **todos** os templates
- Função `showSalvo(msg)` exibe o toast por 2,5 segundos
- **Deve ser chamado em todo save bem-sucedido** — não usar `alert()` para confirmar salvamentos
- Templates que têm: index.html, reajustes.html, dimob.html, informe_anual.html, configuracoes.html, envio_boletos.html

---

## Banner de cota condominial divergente

- Se o valor da cota condominial no boleto do mês for diferente do histórico, exibe banner amarelo
- Campos: `row.cond_alerta`, `row.cond_cota_atual`, `row.cond_cota_hist`
- Banner tem botão fechar (X) e lista todas as unidades afetadas

---

## Bugs já corrigidos — não regredir

### 16. "Novo processamento" não resetava o Mês de Referência
**Problema:** ao carregar a página, `#inMes` é preenchido com o mês atual, mas a restauração automática do localStorage (dados do mês anterior) sobrescrevia esse campo com o mês antigo salvo. Ao clicar em "Novo processamento" pra subir os boletos do mês novo, o campo continuava com o mês antigo (ex: "Junho/2026") em vez do mês atual — e como o valor do campo é enviado como está pro backend (sem validação cruzada com o conteúdo dos PDFs), o processamento novo ficava rotulado com o mês errado, e o Informe Mensal saía com a referência errada.
**Fix:** `novoProcessamento()` agora reseta `#inMes` pro mês atual (mesma lógica da sugestão inicial). **Não remover esse reset.**

### 12. ded{n}_subtrair não salvo no histórico
**Problema:** `salvar_extras` não gravava `ded{n}_subtrair` no historico.json. Ao recarregar no mês seguinte, o campo era undefined → defaultava para `true` (subtrair), virando deduções "+" em "-".  
**Fix:** `salvar_extras` grava `ded{n}_subtrair`; carregamento do histórico restaura o campo diretamente em `row[f"ded{_n}_subtrair"]`. **Não remover.**

### 13. Score de match invertido (cobertura do nome, não do arquivo)
**Problema:** `score_nome_arquivo` media `matches / palavras_do_arquivo`. Filenames com endereço/apto inflavam o denominador → locatário com 2 palavras em arquivo de 8 = 25% → sem match.  
**Fix:** Score agora é `matches / palavras_significativas_do_locatario`. **NÃO reverter.**

### 14. SyntaxError por `const msg` e `let msg` no mesmo escopo
**Problema:** `enviarTodos()` declarava `const msg` e `let msg` na mesma função → SyntaxError que impedia todo o script de carregar (inclusive o toggle do card de cadastro).  
**Fix:** Segunda variável renomeada para `resumo`. **Não reusar o nome `msg` nessa função.**

### 15. Multa/juros ausentes no email de informe
**Problema:** Payload enviado para `/enviar_informe` não incluía `multaBruta`, `jurosBruto`, `taxaSobRec`. Email mostrava repasse correto mas sem discriminar multa/juros.  
**Fix:** Campos adicionados ao payload em ambos os fetchs de informe; `_gerar_html_email` renderiza as linhas. **Manter no payload.**

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

## Múltiplos proprietários por imóvel

### Formato na planilha
- Coluna D (proprietário): `"João Silva, Maria Santos"` — separados por vírgula
- Coluna L (CPF): `"12345678900, 98765432100"` — mesma ordem
- Coluna N (email): `"joao@email.com, maria@email.com"` — mesma ordem

### Comportamento
- `ler_excel_dimob()` detecta vírgula e gera **uma entrada por proprietário** com `aluguel / N`
- `calcular_meses_dimob()` divide `aluguel_antigo` e multa/juros históricos por N
- Informes mensais: `abrirInformes()` split por vírgula → grupo separado por proprietário; `applyDivisor(rd, n)` divide todos os valores monetários
- Cada proprietário recebe seu email individual (posição correspondente na coluna N)
- DIMOB e informes ficam corretos automaticamente — o valor manual/índice gravado na planilha é lido por ambos

### Funções-chave
- `applyDivisor(rd, n)` — divide aluguel, taxa, deduções, multa, juros, repasse por N (index.html)
- `_num_proprietarios`, `_email_individual`, `_proprietario_individual` — campos anotados no row dentro de `abrirInformes()`

---

## Envio de Email DIMOB

- Rota: `POST /api/dimob_enviar_email` — recebe `{proprietario, cpf_proprietario, email_dest, ano, contratos}`
- Função: `_gerar_html_email_dimob()` — gera HTML com todos os imóveis do proprietário, tabela mês a mês
- Email usa URL pública do Railway para logo (não base64)
- Botão **"Enviar por E-mail"** aparece no filtro de proprietário quando um está selecionado
- Botão **"Enviar para Todos"** na barra de ferramentas do DIMOB
- Prioridade de email: 1) emails salvos no modal de informes mensais, 2) coluna N da planilha
- Log gravado em `log_envios.json` com `tipo="dimob"`

---

## Logo nos PDFs e Emails

- Todos os PDFs e emails: `max-width:100%` (ponta a ponta)
- Email informes mensais (`_gerar_html_email`): `max-width:100%`
- Email DIMOB (`_gerar_html_email_dimob`): `max-width:100%`
- Email boletos (`enviar_boleto_locatario`): `max-width:100%`
- **Não usar base64 em emails** — clientes de email bloqueiam. Usar URL pública do Railway
- Fonte do email de boleto: **14pt** (Times New Roman)

## Compatibilidade de email entre clientes

- Layout simples (sem grid/flexbox/media queries) garante renderização consistente em Gmail, Hotmail/Outlook.com e iCloud
- **Imagens externas são bloqueadas por padrão** em todos os clientes (Outlook, Hotmail, iCloud, Gmail) — o destinatário vê o `alt` até clicar "Exibir imagens". Comportamento padrão do mercado, não é bug
- **Outlook desktop (Windows)** pode ignorar `max-width:100%` em imagens — se necessário, trocar por `width="500"` fixo. Atualmente não implementado a pedido
- Texto, fonte e estrutura de parágrafos funcionam em todos os clientes sem problemas

---

## Reajuste Manual

- Campo manual aparece para contratos `RENOVAR` (label "Novo valor do contrato:") **e** `ESTE_MES` (label "Valor manual:", ao lado/abaixo do valor sugerido pelo índice, na mesma célula da tabela)
- `manuais = {}` — dict `num_linha → valor` que prevalece sobre o calculado pelo índice
- Ao aplicar: manual > zerado (0%) > calculado pelo índice
- O valor manual é gravado na coluna F da planilha igual ao reajuste por índice — DIMOB e informes recebem automaticamente
- Lógica de `setManual`/`aplicarReajustes` já era genérica por `num_linha` — não dependia do status, só a exibição do campo no HTML era restrita

---

## Aviso "Confira no site Cálculo Exato"

- Aparece **só em contratos `ESTE_MES`**, na coluna do percentual acumulado (abaixo do valor calculado)
- Backend (`api_calcular_reajustes`) calcula `c['confere_calculo_exato']` (bool) e `c['confere_motivo']` (texto) por contrato
- **Gatilho único:** a variação mensal do **último mês da janela** (o próprio mês do aniversário) veio negativa na série BACEN — usa `historicos_mensal[idx_efetivo]` na chave do mês/ano do aniversário. Vale para qualquer índice (IPCA, INPC ou IGPM)
- **NÃO existe gatilho por "índice ser IGPM"** — foi cogitado e descartado explicitamente; IGPM só dispara o aviso se o último mês vier negativo, igual aos outros índices
- `confere_motivo` é só texto informativo (ex.: "Variação de Junho/2026 negativa") — não é um link, é texto puro
- Sem emoji (segue a regra geral de não usar emojis na interface)
- **Não expandir para FUTURO** sem pedido explícito — decisão consciente de escopo, o "último mês" só importa de fato pros contratos que serão aplicados agora

---

## Parcela IPTU — incremento por mês real

- O histórico (`historico.json`) salva o campo `"saved_month": "AAAA-MM"` em cada entrada
- No processamento, a parcela só é incrementada (+1) se o mês atual for **diferente** do `saved_month`
- Reprocessar dentro do mesmo mês → parcela permanece igual (sobrepõe os dados)
- O botão se chama apenas **"Salvar"**
- Os inputs de parcela (IPTU, IPTU Vaga, Abono) chamam `recalc(idx)` no `oninput` para manter o discriminativo do Total do Boleto sincronizado
- Campo nomeado **"IPTU"** (não "IPTU Apto") — em toda a interface e no Excel exportado

---

## Informe de Rendimento Anual vs DIMOB

### Informe de Rendimento Anual (`/informe_anual`)
- Fluxo direto: calcular → ver → enviar HTML por email diretamente aos proprietários
- "Enviar por E-mail" (proprietário selecionado) e "Enviar para Todos" disponíveis
- Usa a mesma rota `/api/dimob_calcular` do DIMOB
- Template: `templates/informe_anual.html`

### DIMOB (`/dimob`)
- Fluxo com contador: calcular → enviar ao contador → receber PDFs → enviar PDFs aos proprietários
- **Sem** botões de envio direto aos proprietários
- Email do contador: salvo em `config.json["dimob_email_contador"]`, rota `GET/POST /api/dimob_config_contador`
- Upload de PDFs: sistema lê o nome do arquivo, identifica o proprietário por similaridade de nome, busca email no cadastro
- Rota `POST /api/dimob_enviar_pdf_proprietario` — envia PDF como anexo via Resend (base64)
- `_enviar_via_resend` aceita parâmetro `attachments: [{"filename": "...", "content": "<base64>"}]`

---

## Maior índice IPCA/IGPM

- `normalizar_indice()` detecta quando a coluna K contém **ambos** IPCA e IGPM no texto
- Exemplos reconhecidos: `"IPCA/IGPM"`, `"IPCA/IGPM acumulado"`, `"maior IPCA IGPM"`
- Retorna `'MAIOR_IPCA_IGPM'` como `indice_norm`
- No cálculo: busca os dois índices no BACEN, calcula os dois acumulados, aplica o **maior**
- Campo `indice_aplicado` no resultado indica qual foi escolhido ('IPCA' ou 'IGPM')
- Na tabela de reajustes: badge mostra o texto original + `"IGPM (maior)"` em roxo abaixo
- `indices_busca` garante que IPCA e IGPM são sempre buscados quando `MAIOR_IPCA_IGPM` está presente
- **NÃO reverter** a ordem de verificação em `normalizar_indice` — a checagem de ambos deve vir antes das individuais

---

## Confirmação antes de salvar

Todos os botões de salvar têm `confirm()` antes de executar:
- "Salvar" (index.html)
- "Salvar" emails proprietários (modal informes)
- "Salvar Alug. Anteriores" (dimob.html)
- "Salvar E-mails" locatários (envio_boletos.html)
- "Salvar e Testar" API e Email (configuracoes.html)
- "Enviar para Todos" — já tinha confirm nos três lugares (informes, DIMOB, boletos)
- **Envios individuais NÃO pedem confirmação**

---

## Tamanho de fonte

- **Base**: 15px em todos os templates (index, reajustes, dimob, envio_boletos, configuracoes)
- Escala: 10px→12px, 11px→13px, 12px→14px, 13px→15px, 17px→19px
- PDFs (gerados via window.open): usam `pt` — não alterar

---

## O que NÃO adicionar sem pedido explícito

- Ler/processar boletos com demonstrativo (sem cv/vm no nome do arquivo)
- Mover `outros` de volta para `REPASSE_ITENS`
- Filtros "Com Cond." / "Sem Cond." na página de boletos
- Mensagem "Planilha encontrada" na página `/reajustes`
- Botão "Exportar CSV" ou "Exportar Excel" na página de boletos — removido a pedido
- Botão "- 1 Parcela IPTU" — removido a pedido
- Verificação de "vigência mínima de 12 meses"
- Fonte alternativa de dados (IPEA, FGV) para qualquer índice
- Janela de 12 meses (decisão: 13 meses — BACEN Cidadão)
- Coluna "VENCIDO X dias" ou qualquer conceito de atraso
- Alterar a coluna H (data de aniversário) ao aplicar reajuste
- SMTP direto no Railway (portas 465/587 bloqueadas) — usar sempre Resend via HTTPS
- Botão "Informes ao Proprietário" na página de boletos — acesso só pelo card da home
- Emojis em qualquer parte da interface — o sistema não usa emojis

---

## Arquivos principais

```
app.py                        — servidor Flask, toda a lógica backend
templates/home.html           — landing page com cards: Reajustes / Boletos / Envio / DIMOB / Informes
templates/index.html          — página de boletos + informes mensais
templates/reajustes.html      — página de reajuste de aluguel
templates/dimob.html          — página DIMOB (fluxo contador: envio → PDFs → proprietários)
templates/informe_anual.html  — página Informe de Rendimento Anual (envio direto aos proprietários)
templates/envio_boletos.html  — página de envio de boletos por email aos locatários
templates/configuracoes.html  — página de configurações (admin only): API Anthropic + Email/Resend
templates/login.html          — login
static/logo.png               — logo Funchal Imóveis (usada nos PDFs em base64)
.railwayignore                — exclui uploads/ e data/config.json do deploy
Procfile                      — gunicorn, 1 worker, timeout 300s
iniciar.bat                   — inicia o servidor local (janela visível)
iniciar_silencioso.vbs        — inicia o servidor local sem janela (auto-start Windows)
data/config.json              — API key + config SMTP/Resend (criado pela UI)
data/historico.json           — histórico de IPTU, seguros e extras por locatário
data/dimob_historico.json     — histórico de reajustes + multa/juros para o DIMOB
data/log_envios.json          — auditoria de todos os envios de email (informes + boletos)
data/locatarios_emails.json   — emails dos locatários para envio de boletos
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
| `/configuracoes` | GET | **admin** | Página de configurações (API + Email) |
| `/config` | POST | **admin** | Salva API key |
| `/processar` | POST | logado | Processa planilha + PDFs |
| `/salvar_extras` | POST | logado | Salva histórico + multa/juros no DIMOB |
| `/baixar_historico` | GET | logado | Download historico.json |
| `/restaurar_historico` | POST | logado | Restaura historico.json |
| `/configurar_smtp` | POST | **admin** | Salva e testa config email |
| `/get_smtp` | GET | logado | Retorna config email atual |
| `/enviar_informe` | POST | logado | Envia informe por email (Resend) |
| `/api/log_envios` | GET | logado | Retorna log de envios (`?tipo=informe` ou `?tipo=boleto`) |
| `/api/resend_webhook` | POST | público | Recebe eventos de status do Resend |
| `/reajustes` | GET | logado | Página de reajuste |
| `/api/calcular_reajustes` | POST | logado | Calcula reajustes via BACEN |
| `/api/aplicar_reajustes` | POST | logado | Grava novos aluguéis na planilha |
| `/api/baixar_contratos` | GET | logado | Download planilha atualizada |
| `/dimob` | GET | logado | Página DIMOB |
| `/api/dimob_calcular` | POST | logado | Calcula informes anuais |
| `/api/dimob_salvar_historico` | POST | logado | Salva aluguéis anteriores |
| `/api/dimob_exportar` | POST | logado | Exporta Excel DIMOB |
| `/api/dimob_enviar_email` | POST | logado | Envia informe DIMOB por email para um proprietário |
| `/api/dimob_config_contador` | GET/POST | logado | Lê/salva email do contador |
| `/api/dimob_enviar_pdf_proprietario` | POST | logado | Envia PDF (do contador) como anexo ao proprietário |
| `/informe_anual` | GET | logado | Página Informe de Rendimento Anual |
| `/envio_boletos` | GET | logado | Página envio de boletos |
| `/envio_boletos/processar` | POST | logado | Identifica locatários nos PDFs (streaming) |
| `/envio_boletos/enviar` | POST | logado | Envia boleto(s) por email — aceita múltiplos arquivos |
| `/api/locatarios_emails` | GET | logado | Retorna emails cadastrados dos locatários |
| `/api/locatarios_emails/bulk` | POST | logado | Salva emails dos locatários em massa |

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
- Gravados em `dimob_historico.json` quando usuário clica "Salvar"
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

## Decisões de UI — não reverter

### Sem emojis
Toda a interface é sem emojis. Não adicionar emojis em nenhum template HTML, JS ou mensagem visível ao usuário.

### Botão "← Início" em todas as páginas
Todas as 5 páginas (Boletos, Reajustes, DIMOB, Envio de Boletos, Informes) têm botão "← Início" no header que volta à home.

### Página de Configurações (`/configuracoes`)
- Acessível só para admin
- Contém: API Anthropic + configuração de Email/Resend
- Link discreto "Configuracoes" aparece na home (admin only) e no header de Boletos
- Os modais de config foram mantidos no index.html (para compatibilidade JS), mas os botões do header foram removidos

### Cadastro de E-mails (envio_boletos.html)
- Card colapsável — fechado por padrão, abre ao clicar no cabeçalho
- Seta (▼/▲) indica estado aberto/fechado

## localStorage (persistência no browser)
- `boletos_dados` — JSON dos locatários processados
- `boletos_mes` — mês de referência
- `boletos_salvo_em` — "AAAA-MM" do momento do processamento
- Restore automático ao carregar a página; aceita dados do mês atual ou anterior (diff ≤ 1 mês)
