# ✦ Ile Ase Vodun Ogum Ayres

Sistema de gestão para terreiro de Umbanda — consumíveis, acervo, filhos de santo, calendário, entidades/orixás, financeiro e consultas espirituais.

🌐 **Desktop:** [falssp.github.io/terreiro-gestao](https://falssp.github.io/terreiro-gestao)
📱 **Celular:** [falssp.github.io/terreiro-gestao/mobile.html](https://falssp.github.io/terreiro-gestao/mobile.html)
🔐 **Admin:** [falssp.github.io/terreiro-gestao/admin.html](https://falssp.github.io/terreiro-gestao/admin.html)
❓ **Ajuda:** [falssp.github.io/terreiro-gestao/faq.html](https://falssp.github.io/terreiro-gestao/faq.html)

> Acesso dev: ver `dev.md` (privado, nunca commitado).

---

## Funcionalidades

### 🏠 Início (público)
Cards de estoque (OK / Repor / Alerta / Urgente) + eventos do mês + card informativo sobre consultas espirituais (búzios e tarot).

### 📦 Acervo (público — só view)
Roupas, ferramentas e objetos rituais com orixá, entidade, status e localização. Cadastro apenas pelo admin.

### 🕯️ Consumíveis (público — só view)
Estoque com 4 níveis de alerta. Busca de preços no Mercado Livre integrada ao painel admin.

### 📅 Calendário (público)
Giras, Festas, Obrigações, Reuniões, **Consultas de Búzios e Tarot**. Cadastro pelo admin.

### ✦ Orixás & Entidades (público)
**Orixás:** Exu, Ogum, Oxossi, Xangô, Oxum, Iemanjá, Oxalá, Iansã (Oyá), Nanã, Obaluaê/Omolu, Oxumarê, Ossãe

**Entidades / Linhas:** Caboclos, Pretos-Velhos, Exus e Pombagiras, Erês (Crianças), Boiadeiros, Marinheiros, Ciganos

### 👥 Filhos de Santo (privado)
Orixá de cabeça, adjuntó, nação, feitura, aniversário e obrigações. Dados sensíveis -- sem acesso público.

### 💰 Financeiro (privado -- admin)
Lançamento de entradas/saídas com categorias incluindo **Consulta / Atendimento**. Resumo com breakdown por categoria.

### 🛒 Lista de Compras (privado -- admin)
Duas listas separadas:
- **Compras do mês** -- automática pelo nível de estoque
- **Compras pontuais** -- itens avulsos adicionados manualmente
- **Lista completa** -- junta as duas em um só documento

### 📋 Log de Atividades (privado -- admin)
Toda operação de escrita grava: quem fez, quando, aba, campo, valor anterior e novo.

---

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | App desktop |
| `mobile.html` | App celular -- bottom nav |
| `admin.html` | Painel da liderança |
| `faq.html` | Ajuda pública |
| `manifest.json` | Configuração PWA |
| `sw.js` | Cache offline |
| `Terreiro_AppsScript.gs` | Backend Google Apps Script v16.2 |
| `dev.md` | Acesso dev (privado, no .gitignore) |

---

## Perfis de acesso

### Público (sem login)
Acervo, Consumíveis, Calendário, Orixás/Entidades -- somente visualização.

### Filho de Santo (login)
Edita consumíveis e acervo. Vê e edita o próprio perfil. Toda ação registrada no Log.

### Admin -- Pai / Mãe de Santo (login)
Acesso total. Cria usuários com permissões granulares e senha provisória.

### Dev
Acesso via URL com devkey -- ver `dev.md`.

---

## Permissões disponíveis

`acervo_edit` · `consumiveis_edit` · `entidades_edit` · `filhos_view` · `filhos_edit` · `financeiro` · `calendario_edit` · `usuarios` · `configuracoes` · `log`

> Permissões de view (acervo, consumíveis, entidades, calendário) são públicas e não precisam ser configuradas por usuário.

---

## Segurança

- Senha **SHA-256 + salt** -- nunca em texto puro
- **Senha provisória** obriga troca no primeiro acesso
- **3 tentativas** erradas -- bloqueio de 30 minutos
- Token de sessão **em memória apenas** (sem localStorage/sessionStorage) -- expira em 8h
- **Logout automático** por inatividade após 30 minutos
- Acesso dev via URL **não deixa rastro** -- token apenas em memória, URL limpa após entrada

---

## Abas da planilha

**Públicas (via app):** Acervo · Consumíveis · Entidades e Orixás · Calendário

**Privadas:** Filhos de Santo · Financeiro · Lista de Compras · Log · Admins

> A planilha é banco de dados puro. Toda gestão é feita pelo app.

---

## Setup inicial

1. Colar `Terreiro_AppsScript.gs` no Apps Script → rodar `setup()`
2. Reimplantar como Web App -- Executar como: **Eu mesmo** · Acesso: **Qualquer pessoa**
3. Acessar `admin.html` → criar conta do Pai / Mãe de Santo
4. Pai de Santo cria usuários adicionais com permissões granulares

> `setup()` nunca apaga dados existentes. Se a estrutura de uma aba mudar, renomeia a antiga para backup e cria nova vazia.

---

## PWA -- Instalar no celular

**Android (Chrome):** banner automático → "Instalar"

**iPhone (Safari):** botão compartilhar 📤 → Adicionar à Tela de Início

---

## Tecnologias

HTML · CSS · JavaScript (sem frameworks) · Google Apps Script · Google Sheets · PWA · Mercado Livre API

---

*v16.2 -- 01/08/2026*
