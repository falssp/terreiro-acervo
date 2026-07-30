# ✦ Ile Ase Vodun Ogum Ayres

Sistema de gestão para terreiro de Umbanda — consumíveis, acervo, filhos de santo, entidades/orixás e financeiro.

🌐 **Desktop:** [falssp.github.io/terreiro-gestao](https://falssp.github.io/terreiro-gestao)
📱 **Celular:** [falssp.github.io/terreiro-gestao/mobile.html](https://falssp.github.io/terreiro-gestao/mobile.html)
❓ **Ajuda:** [falssp.github.io/terreiro-gestao/faq.html](https://falssp.github.io/terreiro-gestao/faq.html)

---

## Funcionalidades

### 📦 Acervo
Cadastro de roupas, indumentárias, ferramentas e objetos rituais com orixá associado, status, localização e foto.

### 🕯️ Consumíveis
Controle de estoque com 4 níveis de alerta:
- ✅ **OK** — ≥ 75% do mínimo
- ⚠️ **Repor** — 25 a 74%
- 🔴 **Alerta** — 5 a 24%
- 🚨 **Urgente / Zero** — menos de 5%

Categorias: Velas · Bebidas · Fumo · Ervas e Defumação · Alimentos · Flores e Naturais

Busca de preços no **Mercado Livre** direto do app.

### 👥 Filhos de Santo
Cadastro com orixá de cabeça, adjuntó, nação, data de feitura, aniversário e controle de obrigações.
> O orixá de cabeça **não aparece** no app público — segredo de santo. Fica registrado apenas na planilha.

### 🥁 Entidades e Orixás
Referência com 13 orixás/entidades pré-cadastrados: oferendas, cores, saudações e datas de festa.

### 💰 Financeiro
Lançamento de entradas e saídas com categorias, responsável e comprovante. Painel de saldo total e do mês.

### 📅 Início
Painel do mês com aniversariantes, obrigações e festas de orixás. Carrega automaticamente ao abrir.

---

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | App desktop (tabs no topo) |
| `mobile.html` | App celular (bottom nav, otimizado para toque) |
| `admin.html` | Painel da liderança — acesso restrito por login |
| `faq.html` | Ajuda pública (3 perguntas essenciais) |
| `manifest.json` | Configuração PWA |
| `sw.js` | Cache offline (Service Worker) |
| `icon-192.png` / `icon-512.png` | Ícones para tela inicial |
| `Terreiro_AppsScript.gs` | Backend Google Apps Script (referência) |

---

## Sistema de login (admin.html)

- Login por **e-mail + senha individual**
- Senha armazenada como **SHA-256 + salt** — nunca em texto puro
- **3 tentativas** erradas → bloqueio de 30 minutos automático
- Desbloqueio pelo próprio painel (sem mexer na planilha)
- Token de sessão com expiração de 8 horas
- **Primeiro acesso:** detecta aba Admins vazia → tela de criação do Pai de Santo
- Permissões configuráveis por usuário: `estoque` · `datas` · `obrigacoes` · `financeiro` · `mailing` · `filhos` · `configuracoes`

---

## Planilha

→ [Acervo do Terreiro — Cadastro de Itens](https://docs.google.com/spreadsheets/d/1zzP0TgsT85omd2MyIGLU-B2ONRMwI6aQkgQO4DroIFc/edit)

Abas: **Acervo · Consumíveis · Filhos de Santo · Entidades e Orixás · Financeiro · Admins**

Menu **🏛️ Terreiro** na planilha: recalcular estoque, lista de compras, datas do mês, obrigações.

> A planilha é usada como banco de dados. Ações de gestão (desbloquear admin, criar usuários, etc.) são feitas pelo app — não editando a planilha diretamente.

---

## Tecnologias

- HTML / CSS / JavaScript — sem frameworks
- Google Apps Script — backend e API
- Google Sheets — banco de dados
- PWA — instalável no celular (Android e iPhone)
- Mercado Livre API — busca de preços

---

## Instalar no celular

**Android (Chrome):** banner aparece automaticamente → "Instalar". Se não: 3 pontos → Adicionar à tela inicial.

**iPhone (Safari):** botão compartilhar 📤 → Adicionar à Tela de Início. Funciona apenas no Safari.

---

*Desenvolvido para uso interno do terreiro.*
