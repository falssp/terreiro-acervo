# ✦ Acervo do Terreiro

Sistema de gestão para terreiro de Umbanda — acervo de itens, consumíveis, filhos de santo, entidades/orixás e financeiro.

🌐 **App:** [falssp.github.io/terreiro-acervo](https://falssp.github.io/terreiro-acervo)
❓ **Ajuda:** [falssp.github.io/terreiro-acervo/faq.html](https://falssp.github.io/terreiro-acervo/faq.html)

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
> O orixá de cabeça **não aparece** na listagem pública do app — fica registrado apenas na planilha (segredo de santo).

### 🥁 Entidades e Orixás
Referência com 13 orixás/entidades pré-cadastrados: oferendas, cores, saudações e datas de festa.

### 💰 Financeiro
Lançamento de entradas e saídas com categorias, responsável e comprovante. Painel de saldo total e do mês.

### 📅 Início
Painel do mês com aniversariantes, obrigações e festas de orixás.

---

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | App web completo (PWA) |
| `manifest.json` | Configuração do app instalado |
| `sw.js` | Cache offline (Service Worker) |
| `icon-192.png` / `icon-512.png` | Ícones para tela inicial |
| `faq.html` | Página de ajuda |
| `Terreiro_AppsScript.gs` | Backend Google Apps Script (referência) |

---

## Tecnologias

- HTML / CSS / JavaScript puro — sem frameworks
- Google Apps Script — backend e API
- Google Sheets — banco de dados
- PWA — instalável no celular (Android e iPhone)
- Mercado Livre API — busca de preços (gratuita)

---

## Planilha

→ [Acervo do Terreiro — Cadastro de Itens](https://docs.google.com/spreadsheets/d/1zzP0TgsT85omd2MyIGLU-B2ONRMwI6aQkgQO4DroIFc/edit)

Abas: **Acervo · Consumíveis · Filhos de Santo · Entidades e Orixás · Financeiro**

Menu **🏛️ Terreiro** na planilha com atalhos: recalcular estoque, lista de compras, datas do mês, obrigações próximas.

---

## Instalar no celular

**Android (Chrome):** banner aparece automaticamente → "Instalar". Se não aparecer: 3 pontos → Adicionar à tela inicial.

**iPhone (Safari):** botão compartilhar 📤 → Adicionar à Tela de Início.

---

*Desenvolvido para uso interno do terreiro.*
