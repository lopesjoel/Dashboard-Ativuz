# Como instalar o app de Vistoria

Esse app roda dentro da sua conta Google, de graça, usando o **Google Apps Script**.
Ele não precisa de servidor, hospedagem ou custo mensal.

Esta versão já está adaptada ao modelo "Anexo I do Contrato de Locação — Ordem de
Serviço e Vistoria" da ATIVUZ: dados do cliente/veículo, checklist de acessórios
(S/N/A), luzes do painel, fotos organizadas por área do carro, observações,
descrição dos sintomas e assinatura do cliente + do responsável.

**Esta versão substitui a página `/vistoria` do Dashboard.** Toda vistoria
enviada aqui também é gravada no Dashboard (Supabase): gera o mesmo `.docx`
oficial e aparece no Histórico de Vistorias, com vínculo ao contrato quando
selecionado. A cópia no Drive/planilha continua sendo feita também, como
backup.

## Passo 1 — Criar o projeto
1. Acesse **script.google.com** e clique em **Novo projeto**.
2. Dê um nome, por exemplo "Vistoria de Veículos".
3. Apague o conteúdo padrão do arquivo `Code.gs` e cole o conteúdo do arquivo `Code.gs` deste pacote.
4. Clique no `+` ao lado de "Arquivos" → **HTML** → nomeie exatamente `Index` (sem `.html`) → cole o conteúdo do arquivo `Index.html` deste pacote.

## Passo 2 — Pasta dos contratos (já configurado)
O `PARENT_FOLDER_ID` já está preenchido no `Code.gs` apontando pra pasta **"2. Contratos"** do Drive da Ativuz. Essa pasta tem uma subpasta por ano (`! 2023`, `! 2024`, `! 2025`, `! 2026`...) e, dentro de cada ano, uma pasta por contrato no padrão `NN/ANO_PLACA_Nome` (ex: `12/2026_QGT6I05_Matheus Sousa`).

Quando o vistoriador seleciona um contrato no formulário, o app procura — em todas as pastas de ano — uma pasta cujo nome contenha a **placa** do contrato, e salva as fotos/assinaturas dentro dela (numa subpasta `Fotos Vistoria`). Se não encontrar (ex: contrato novo, pasta ainda não criada por quem organiza o Drive), o app cria uma pasta nova dentro da pasta do ano atual — **nunca apaga ou move nada que já existia**.

Se a estrutura de pastas mudar de lugar no futuro, só é preciso trocar o ID em:
```
const PARENT_FOLDER_ID = '1srRiEPkZOGaVQoF3MQcMSy_RAkGdXkfv';
```

## Passo 2.5 — Conectar ao Dashboard Ativuz
Para a vistoria aparecer no Histórico de Vistorias do Dashboard (e gerar o `.docx` oficial automaticamente):
1. No `Code.gs`, preencha:
   ```
   const DASHBOARD_API_URL = 'https://SEU-DOMINIO-DO-DASHBOARD';   // sem barra no final
   const DASHBOARD_API_TOKEN = 'MESMO-VALOR-DE-VISTORIA_API_TOKEN-NO-.ENV-DO-DASHBOARD';
   ```
2. O token é um segredo — só ele impede que qualquer pessoa na internet grave vistorias falsas no seu banco de dados. Não compartilhe fora da equipe técnica.
3. Se o Dashboard estiver fora do ar no momento do envio, a vistoria continua sendo salva normalmente no Drive/planilha — só a parte do Dashboard falha, e isso aparece avisado na tela e na coluna "Dashboard: status" da planilha de log.

## Passo 3 — Publicar como Web App
1. No editor do Apps Script, clique em **Implantar** → **Nova implantação**.
2. Em "Tipo", escolha **App da Web**.
3. Em "Executar como": **Eu (sua conta)**.
4. Em "Quem pode acessar": escolha conforme sua necessidade —
   - **Qualquer pessoa com o link** → motoristas de fora da organização também acessam sem login Google.
   - **Qualquer pessoa na [sua organização]** → exige estar logado com e-mail da empresa (mais seguro).
5. Clique em **Implantar**, autorize as permissões (vai pedir acesso ao Drive e Planilhas — é esperado, é o que faz o app funcionar).
6. Copie o **link do App da Web** gerado. Esse é o link que você vai usar/compartilhar (pode salvar como atalho na tela inicial do celular).

## Passo 4 — Testar
1. Abra o link, escolha a etapa (Entrada ou Saída), selecione um contrato ativo (se aparecer na lista) ou digite o cliente manualmente, tire fotos, preencha o checklist, assine e envie.
2. Confira se a pasta `Vistoria Entrada dd-mm-aaaa_HH-mm` (ou `Vistoria Saida ...`) apareceu dentro de `Fotos Vistoria` na pasta do motorista.
3. Uma planilha chamada **"Registro de Vistorias"** é criada automaticamente no seu Drive na primeira vez — é o seu histórico central de todas as vistorias (agora com colunas de Etapa e status do envio pro Dashboard).
4. No Dashboard, confira em **Histórico de Vistorias** se a vistoria apareceu e se o `.docx` foi gerado corretamente.

## Atualizando depois de publicado
Sempre que editar `Code.gs` ou `Index.html`, é preciso ir em **Implantar → Gerenciar implantações → editar (ícone de lápis) → Nova versão → Implantar** para o link já publicado refletir as mudanças.

## Preenchimento automático a partir do contrato
- A seleção de contrato é **obrigatória**: ao escolher um contrato na lista, o cliente/veículo/placa já vêm direto do Dashboard, e o app busca telefone/endereço/cor/ano/chassi/número do motor direto da planilha `planilhas/DADOS_CLIENTES_CONS.xlsx` do Dashboard (rota `/api/contrato/dados`, buscada pela placa) — sem OCR, sem depender de nenhum arquivo dentro da pasta do contrato no Drive.
- Os campos são preenchidos automaticamente, mas continuam **editáveis** — sempre confira antes de enviar, principalmente na primeira vez. O botão "Buscar dados do contrato novamente" repete a busca se algo não tiver vindo certo.
- Se a placa não for encontrada nessa planilha (ex: contrato muito novo, ainda não cadastrado lá), o app avisa na tela e os campos continuam em branco pra preenchimento manual, sem travar o formulário.
- A pasta do Drive onde ficam as fotos/assinaturas continua sendo localizada pela placa (Passo 2) — isso não muda, só o preenchimento automático dos campos é que não depende mais de nenhum documento na pasta.

## Como funciona a integração com o Dashboard
- **Etapa**: escolha Entrada (retirada do carro) ou Saída (devolução) no topo do formulário. O Dashboard trata cada vistoria como um par entrada+saída do mesmo contrato — a etapa de Saída localiza automaticamente a entrada mais recente do contrato selecionado e completa o mesmo registro, gerando o `.docx` final com os dois lados.
- **Contrato**: o formulário busca a lista de contratos ativos direto do Dashboard e exige que um seja selecionado antes de enviar. Selecionar um contrato preenche cliente/veículo/placa, vincula a vistoria a ele (o que faz aparecer corretamente ligada ao contrato certo no Histórico) e é o que permite localizar a pasta certa no Drive pela placa. Se o Dashboard estiver fora do ar e a lista não carregar, o formulário fica bloqueado até recarregar — não há mais preenchimento manual sem contrato.
- **Assinaturas**: são enviadas e guardadas no Storage do Dashboard, mas por enquanto não aparecem dentro do `.docx` gerado (o modelo de documento ainda não tem esse campo) — ficam disponíveis como imagem separada para conferência.
- **Se a chamada pro Dashboard falhar** (fora do ar, token errado, etc.): a vistoria continua sendo salva no Drive/planilha normalmente, e o app avisa na tela que só a parte do Dashboard não foi. Não é preciso preencher tudo de novo — o mais simples nesse caso é reenviar depois manualmente pelo Dashboard, ou me avisar que eu ajudo a reprocessar.

## Personalizações rápidas
- **Cores da empresa**: no topo do `Index.html`, dentro de `:root`, troque `--cor-primaria` pela cor da sua marca (ex: o azul da ATIVUZ).
- **Logo**: troque o bloco `<div class="logo-placeholder">LOGO</div>` por `<img src="URL_PUBLICA_DO_SEU_LOGO" style="height:34px">`.
- **Itens de acessórios**: edite a lista `ACCESSORY_ITEMS` no `Code.gs`. **Atenção**: a ordem precisa bater com `_CHAVES_ACC` no `app.py` do Dashboard — se reordenar aqui sem ajustar lá, os status (S/N/A) vão gravar no item errado.
- **Categorias de foto**: edite a lista `PHOTO_CATEGORIES` no `Code.gs` (hoje: os 12 ângulos que o `.docx` do Dashboard usa — frontal, traseira, lateral direita/esquerda, painel, hodômetro, estepe, teto, motor, porta-malas, dano 1, dano 2). Cada item é `{ key, label }` — pode mudar o `label` (texto exibido) livremente, mas **não mude o `key`**, ele precisa bater com `ANGULOS_FOTO` em `services/gerar_vistoria_entrada_saida.py` no Dashboard.

## O que mudou em relação ao papel
- O desenho do carro para marcar avarias à mão foi substituído por **fotos organizadas por área do veículo** — mais rápido de preencher no celular e mais fácil de auditar depois (cada foto já entra com o nome da categoria).
- As **luzes do painel acesas** não são mais marcadas item a item: a foto da categoria "Painel / Hodômetro" (tirada com o veículo ligado) já registra isso visualmente, com evidência mais confiável do que uma marcação manual.
- Toda foto recebe automaticamente uma **marca d'água com data e hora** (canto inferior direito), carimbada no momento em que a foto é selecionada no formulário — funciona tanto para fotos tiradas na hora quanto vindas da galeria.
- As duas assinaturas (cliente e responsável) são coletadas na tela, com o dedo ou mouse.
- Cada vistoria vira uma linha na planilha "Registro de Vistorias", com link direto pra pasta de fotos daquela vistoria específica.

## Sobre a marca d'água de data/hora
- A marca usa a **data/hora do próprio celular/navegador** de quem está preenchendo — não é um carimbo criptográfico à prova de fraude, é o mesmo tipo de marca que qualquer app de câmera com timestamp usa (serve como evidência visual, não como prova legal inviolável).
- Se quiser um nível mais forte de confiabilidade (por exemplo, provar que a foto não foi enviada de uma foto antiga tirada por fora do app), dá pra evoluir depois adicionando também **geolocalização** na marca, já que ela é obtida por GPS do aparelho no momento do envio — me avise se quiser essa camada extra.
