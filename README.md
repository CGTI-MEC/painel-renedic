# Painel Renedic — adesão das redes de ensino

Painel de acompanhamento da adesão das redes municipais e estaduais à Renedic,
com dados do relatório **Enec_Status_Adesao_Detalhado** do SIMEC.

---

## Parte 1 — Configuração inicial (fazer uma única vez)

### 1.1 Criar o repositório no GitHub

1. Entre em <https://github.com> com a mesma conta do Painel Enec.
2. No canto superior direito, clique em **+** → **New repository**.
3. Preencha:
   - **Repository name:** `painel-renedic`
   - **Visibility:** **Public** (o GitHub Pages gratuito exige repositório público)
   - **Não** marque "Add a README", ".gitignore" nem "license" (os arquivos já estão prontos).
4. Clique em **Create repository**. Deixe essa página aberta; você vai usar o endereço que aparece nela
   (algo como `https://github.com/SEU-USUARIO/painel-renedic.git`).

### 1.2 Montar a pasta no seu computador

1. Crie uma pasta, por exemplo `Documentos\painel-renedic`.
2. Copie para ela todos os arquivos entregues:

```
painel-renedic\
  index.html          ← o painel
  processar.py        ← lê o relatório e gera os dados
  atualizar.bat       ← processa e publica (duplo clique)
  relatorio.txt       ← mensagem, prazos e chave de admin (NÃO vai para o GitHub)
  .gitignore          ← impede que pastas sensíveis subam para o GitHub
  README.md           ← este guia
  brutos\             ← crie: aqui você larga o relatório do SIMEC
  referencia\         ← crie: aqui fica a base do Censo
```

3. Coloque o arquivo `Enec_escolas_2026_base_censo25.xlsx` dentro de `referencia\`.
4. Coloque o relatório do SIMEC (`Enec_Status_Adesao_Detalhado_DD-MM-AAAA_HH_MM.xls`) dentro de `brutos\`.

> As pastas `brutos\` e `referencia\` e o `relatorio.txt` **nunca** são enviados ao GitHub
> (o `.gitignore` bloqueia). Os nomes dos responsáveis que vêm no relatório são descartados
> pelo script antes de gerar qualquer arquivo do painel.

### 1.3 Definir a chave de admin

Abra o `relatorio.txt` no Bloco de Notas e troque `troque-esta-chave` por uma chave sua
(sem espaços, por exemplo `renedic-equipe-2026x`). Salve.

### 1.4 Instalar os complementos do Python (uma vez)

Abra o **Prompt de Comando** e rode:

```
py -3.11 -m pip install pandas openpyxl lxml xlrd
```

(Se o Painel Enec já roda neste computador, provavelmente já está tudo instalado; rodar de novo não faz mal.)

### 1.5 Ligar a pasta ao GitHub e fazer a primeira publicação

No Explorador de Arquivos, abra a pasta `painel-renedic`, clique na barra de endereço,
digite `cmd` e tecle Enter. No Prompt que abrir, rode **um comando por vez**
(troque `SEU-USUARIO` pelo seu usuário do GitHub):

```
git init -b main
git remote add origin https://github.com/SEU-USUARIO/painel-renedic.git
py -3.11 processar.py
git add .gitignore README.md index.html processar.py atualizar.bat saida
git commit -m "Primeira versao do Painel Renedic"
git push -u origin main
```

Se o GitHub pedir login, uma janela do navegador abre para você autorizar. É o mesmo processo do Painel Enec.

> Confira depois, na página do repositório no GitHub, que **não** aparecem as pastas
> `brutos`, `referencia` nem o arquivo `relatorio.txt`.

### 1.6 Ativar o site (GitHub Pages)

1. No repositório, vá em **Settings** → **Pages** (menu da esquerda).
2. Em **Build and deployment** → **Source**, escolha **Deploy from a branch**.
3. Em **Branch**, escolha **main** e a pasta **/ (root)**. Clique em **Save**.
4. Aguarde 1 a 2 minutos e recarregue a página. Aparece o endereço do painel:
   `https://SEU-USUARIO.github.io/painel-renedic/`

### 1.7 Os dois links

| Para quem | Link |
|---|---|
| Articuladores | `https://SEU-USUARIO.github.io/painel-renedic/` |
| Equipe (admin) | `https://SEU-USUARIO.github.io/painel-renedic/?admin=SUA-CHAVE` |

A visão da equipe mostra a aba **Redes estaduais**, os filtros do Censo (Capital, Marajó,
quilombola, indígena, campo, rural, assentamento, comunidade tradicional), a lista completa
sem precisar de filtro e a exportação da base inteira (com as colunas do Censo).

> A chave de admin esconde a visão da equipe, mas não é uma senha forte: os dados do painel
> (status por rede, sem nomes de pessoas) ficam em arquivos públicos do repositório.
> Não coloque no painel nada que não possa ser público.

---

## Parte 2 — Uso no dia a dia

1. Baixe o relatório no SIMEC e salve em `brutos\` **sem renomear**
   (o nome traz a data e hora usadas no painel).
2. Dê **duplo clique** no `atualizar.bat`.
3. Em 1 a 2 minutos o painel está atualizado. Quem estiver com o painel aberto recebe os
   dados novos sozinho em até 10 minutos.

Pode deixar os relatórios antigos na pasta `brutos\`: o script sempre usa o mais recente.

### Prorrogação de prazo ou nova mensagem

Edite o `relatorio.txt` (por exemplo, `PRAZO_MUNICIPAL: 27/11/2026`), salve e dê duplo clique no
`atualizar.bat`. Não precisa de relatório novo.

### Curva de evolução

O arquivo `saida\envios.json` guarda o dia em que cada rede apareceu pela primeira vez
como "Em análise". **Não apague esse arquivo**: ele é o histórico da curva.
- No primeiro processamento, as redes já em análise usam a data do termo aceito.
- Daí em diante, vale a data do relatório em que a rede apareceu em análise.
- Várias atualizações no mesmo dia contam como um único dia.

---

## Problemas comuns

| Mensagem | O que fazer |
|---|---|
| `Status desconhecido(s) no relatório: X` | O SIMEC criou um status novo. Abra o `processar.py`, acrescente o texto (sem acento, minúsculo) no `STATUS_MAP` e rode de novo. |
| `O relatório em brutos/ é MAIS ANTIGO que o já publicado` | Você colocou um arquivo antigo com data mais nova no computador. Baixe o relatório atual. |
| `Esperado ~5.569 redes municipais e vieram N` | O relatório pode ter vindo incompleto. Baixe de novo. |
| `Coluna não encontrada` | O SIMEC mudou o layout do relatório. Envie o cabeçalho (sem os nomes) para ajustar o script. |
| Erro no `git push` | Verifique a internet e se está logado no GitHub. |
| Painel não abre os dados | Confira, no GitHub, se a pasta `saida` tem os 4 arquivos `.json`. |

---

## O que cada arquivo gera

| Arquivo em `saida\` | Conteúdo |
|---|---|
| `redes.json` | Uma linha por rede: UF, município, IBGE, status, data do termo aceito, marcações do Censo |
| `envios.json` | Histórico da curva de evolução |
| `config.json` | Mensagem, prazos e o resumo (hash) da chave de admin |
| `meta.json` | Data e hora do relatório usado |

A coluna "Responsável" do relatório nunca é gravada. Brasília (DF) e Fernando de Noronha (PE)
são removidos das redes municipais, porque não têm rede municipal.
