# ✦ Ile Ase Vodun Ogum Ayres

Sistema de gestão para terreiro de Umbanda — consumíveis, acervo, filhos de santo, calendário, entidades/orixás e financeiro.

🌐 **Desktop:** [falssp.github.io/terreiro-gestao](https://falssp.github.io/terreiro-gestao)
📱 **Celular:** [falssp.github.io/terreiro-gestao/mobile.html](https://falssp.github.io/terreiro-gestao/mobile.html)
🔐 **Admin:** [falssp.github.io/terreiro-gestao/admin.html](https://falssp.github.io/terreiro-gestao/admin.html)
❓ **Ajuda:** [falssp.github.io/terreiro-gestao/faq.html](https://falssp.github.io/terreiro-gestao/faq.html)

---

## Funcionalidades

### 🏠 Início (público)

Painel do mês com eventos do calendário e festas de orixás. Cards clicáveis de estoque (OK / Repor / Alerta / Urgente) levam direto para os consumíveis filtrados.

### 📦 Acervo (público — só view)

Cadastro de roupas, indumentárias, ferramentas e objetos rituais com orixá de cabeça, entidade / linha, status, localização e foto. Cadastro via app (requer token).

### 🕯️ Consumíveis (público — view; login para editar)

Controle de estoque com 4 níveis de alerta. Clique no card → modal com dados → botão "Atualizar estoque".

Níveis:
- ✅ **OK** — ≥ 75% do mínimo
- ⚠️ **Repor** — 25 a 74%
- 🔴 **Alerta** — 5 a 24%
- 🚨 **Urgente / Zero** — menos de 5%

Categorias: Alimentos · Bebidas · Ervas e Defumação · Flores e Naturais · Fumo · Velas · Outro

Busca de preços no **Mercado Livre** direto do app.

### 📅 Calendário (público)

Eventos do terreiro: Giras, Festas, Obrigações, Reuniões. Cadastro pelo painel admin.

### ✦ Orixás e Entidades (público)

Referência com orixás/entidades pré-cadastrados: oferendas, cores, saudações e datas de festa.

### 👥 Filhos de Santo (privado — login obrigatório)

Cadastro com orixá de cabeça, adjuntó, nação, data de feitura, aniversário e controle de obrigações. Dados sensíveis — sem acesso público.

### 💰 Financeiro (privado — admin)

Lançamento de entradas e saídas com categorias. Painel de saldo acessível apenas para admins com permissão `financeiro`.

### 📋 Log de Atividades (privado — admin)

Toda operação de escrita grava automaticamente: quem fez, quando, em qual aba, campo alterado, valor anterior e novo.

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
| `Terreiro_AppsScript.gs` | Backend Google Apps Script v16 |

---

## Acesso e perfis

### Público (sem login)
Acervo, Consumíveis, Calendário, Orixás/Entidades — somente visualização.

### Filho de Santo (login)
Edita consumíveis e acervo. Vê e edita o próprio perfil em Filhos de Santo. Toda ação é registrada no Log.

### Admin — Pai / Mãe de Santo (login)
Acesso total. Cria usuários com permissões granulares e senha provisória (obriga troca no primeiro acesso).

### Dev
Acesso via `devkey` na URL — sem cadastro na planilha.

---

## Permissões disponíveis

`acervo_view` · `acervo_edit` · `consumiveis_view` · `consumiveis_edit` · `entidades_view` · `entidades_edit` · `filhos_view` · `filhos_edit` · `financeiro` · `calendario_view` · `calendario_edit` · `usuarios` · `configuracoes` · `log`

---

## Segurança

- Login por **e-mail + senha individual**
- Senha armazenada como **SHA-256 + salt** — nunca em texto puro
- **Senha provisória** no cadastro de novo usuário → troca obrigatória no primeiro acesso
- **3 tentativas** erradas → bloqueio automático de 30 minutos
- Desbloqueio pelo painel (sem editar a planilha)
- Token de sessão em memória apenas (sem localStorage/sessionStorage) — expira em 8h
- **Logout automático por inatividade** após 30 minutos sem interação

---

## Planilha

→ [Gestão do Terreiro](https://docs.google.com/spreadsheets/d/1zzP0TgsT85omd2MyIGLU-B2ONRMwI6aQkgQO4DroIFc/edit)

Abas: **Acervo · Consumíveis · Entidades e Orixás · Calendário · Filhos de Santo · Financeiro · Log · Admins**

> A planilha é banco de dados puro. Toda gestão é feita pelo app — não edite a planilha diretamente.

Menu **🏛️ Gestão do Terreiro** na planilha: recalcular estoque, lista de compras, datas do mês, obrigações próximas.

---

## Setup inicial

1. Colar `Terreiro_AppsScript.gs` no Apps Script → rodar `setup()`
2. Reimplantar como Web App — Executar como: **Eu mesmo** · Acesso: **Qualquer pessoa**
3. Acessar `admin.html` → tela de criação do Pai / Mãe de Santo
4. Pai de Santo cria usuários adicionais pelo painel com permissões granulares

---

## Instalar no celular (PWA)

**Android (Chrome):** banner aparece automaticamente → "Instalar". Se não: 3 pontos → Adicionar à tela inicial.

**iPhone (Safari):** botão compartilhar 📤 → Adicionar à Tela de Início. Funciona apenas no Safari.

---

## Tecnologias

HTML · CSS · JavaScript — sem frameworks · Google Apps Script · Google Sheets · PWA · Mercado Livre API

---

*v16 — 01/08/2026 · Desenvolvido para uso interno do terreiro.*
