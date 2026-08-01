# ✦ Ile Ase Vodun Ogum Ayres

Sistema de gestão para terreiro de Umbanda — consumíveis, acervo, filhos de santo, entidades/orixás e financeiro.

🌐 **Desktop:** [falssp.github.io/terreiro-gestao](https://falssp.github.io/terreiro-gestao)
📱 **Celular:** [falssp.github.io/terreiro-gestao/mobile.html](https://falssp.github.io/terreiro-gestao/mobile.html)
🔐 **Admin:** [falssp.github.io/terreiro-gestao/admin.html](https://falssp.github.io/terreiro-gestao/admin.html)
❓ **Ajuda:** [falssp.github.io/terreiro-gestao/faq.html](https://falssp.github.io/terreiro-gestao/faq.html)

---

## Funcionalidades

### 🏠 Início
Painel do mês com aniversariantes, obrigações e festas de orixás. Cards clicáveis de estoque (OK / Repor / Alerta / Urgente) levam direto para os consumíveis filtrados.

### 📦 Acervo
Cadastro de roupas, indumentárias, ferramentas e objetos rituais com orixá, entidade, status, localização e foto.

### 🕯️ Consumíveis
Controle de estoque com 4 níveis de alerta. Clique no card → modal com dados → botão "Atualizar estoque" preenche o formulário automaticamente.

Níveis:
- ✅ **OK** — ≥ 75% do mínimo
- ⚠️ **Repor** — 25 a 74%
- 🔴 **Alerta** — 5 a 24%
- 🚨 **Urgente / Zero** — menos de 5%

Categorias: Alimentos · Bebidas · Ervas e Defumação · Flores e Naturais · Fumo · Velas · Outro

Busca de preços no **Mercado Livre** direto do app (aba Cadastrar).

### 👥 Filhos de Santo
Cadastro com orixá de cabeça, adjuntó, nação, data de feitura, aniversário e controle de obrigações. O orixá de cabeça **não aparece** no app público — segredo de santo.

### 🥁 Entidades e Orixás
Referência com orixás/entidades pré-cadastrados: oferendas, cores, saudações e datas de festa.

### 💰 Financeiro
Lançamento de entradas e saídas com categorias e responsável. Painel de saldo acessível apenas no admin.

---

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | App desktop — responsivo, tabs no topo |
| `mobile.html` | App celular — bottom nav, otimizado para toque |
| `admin.html` | Painel da liderança — acesso por login |
| `faq.html` | Ajuda pública |
| `manifest.json` | Configuração PWA |
| `sw.js` | Cache offline |
| `Terreiro_AppsScript.gs` | Backend Google Apps Script |

---

## Sistema de login (admin.html)

- Login por **e-mail + senha individual**
- Senha armazenada como **SHA-256 + salt** — nunca em texto puro
- **3 tentativas** erradas → bloqueio de 30 minutos automático
- Desbloqueio pelo painel (sem mexer na planilha)
- Token de sessão com expiração de 8 horas
- **Primeiro acesso:** detecta aba Admins vazia → tela de criação do Pai de Santo
- Permissões por usuário: `estoque` · `datas` · `obrigacoes` · `financeiro` · `mailing` · `filhos` · `configuracoes`

---

## Planilha

→ [Gestão do Terreiro](https://docs.google.com/spreadsheets/d/1zzP0TgsT85omd2MyIGLU-B2ONRMwI6aQkgQO4DroIFc/edit)

Abas: **Acervo · Consumíveis · Filhos de Santo · Entidades e Orixás · Financeiro · Admins**

Menu **🏛️ Terreiro** na planilha: recalcular estoque, lista de compras, datas do mês, obrigações.

> A planilha é banco de dados. Ações de gestão (criar usuários, desbloquear admin, atualizar estoque) são feitas pelo app — não editando a planilha diretamente.

---

## Setup inicial

1. Colar `Terreiro_AppsScript.gs` no Apps Script → rodar `setup()`
2. Reimplantar como nova versão (mantém a mesma URL)
3. Acessar `admin.html` → tela de criação do Pai de Santo
4. Pai de Santo cria admins adicionais pelo painel

---

## Instalar no celular (PWA)

**Android (Chrome):** banner aparece automaticamente → "Instalar". Se não: 3 pontos → Adicionar à tela inicial.

**iPhone (Safari):** botão compartilhar 📤 → Adicionar à Tela de Início. Funciona apenas no Safari.

---

## Tecnologias

HTML · CSS · JavaScript — sem frameworks · Google Apps Script · Google Sheets · PWA · Mercado Livre API

---

*Desenvolvido para uso interno do terreiro.*
