# Como instalar o app de Vistoria

Esse app roda dentro da sua conta Google, de graça, usando o **Google Apps Script**.
Ele não precisa de servidor, hospedagem ou custo mensal.

Esta versão já está adaptada ao modelo "Anexo I do Contrato de Locação — Ordem de
Serviço e Vistoria" da ATIVUZ: dados do cliente/veículo, checklist de acessórios
(S/N/A), luzes do painel, fotos organizadas por área do carro, observações,
descrição dos sintomas e assinatura do cliente + do responsável.

## Passo 1 — Criar o projeto
1. Acesse **script.google.com** e clique em **Novo projeto**.
2. Dê um nome, por exemplo "Vistoria de Veículos".
3. Apague o conteúdo padrão do arquivo `Code.gs` e cole o conteúdo do arquivo `Code.gs` deste pacote.
4. Clique no `+` ao lado de "Arquivos" → **HTML** → nomeie exatamente `Index` (sem `.html`) → cole o conteúdo do arquivo `Index.html` deste pacote.

## Passo 2 — Configurar a pasta dos clientes/motoristas
1. No Drive, entre na pasta que contém as pastas de cada cliente/motorista (a pasta "mãe").
2. Copie o ID da pasta pela URL: `drive.google.com/drive/folders/**ESTE-TRECHO**`.
3. No `Code.gs`, cole esse ID na linha:
   ```
   const PARENT_FOLDER_ID = 'COLOQUE_AQUI_O_ID_DA_PASTA_MAE';
   ```
4. Se o cliente já tiver pasta com esse nome exato, o app usa ela. Se digitar um nome novo, o app **cria a pasta automaticamente** dentro da pasta mãe — não precisa mais criar pasta na mão para cliente novo.

## Passo 3 — Ativar o serviço de OCR (necessário para ler o contrato)
O preenchimento automático a partir do contrato em PDF usa OCR do Google Drive, que é um "serviço avançado" e precisa ser ligado uma vez só:
1. No editor do Apps Script, clique no `+` ao lado de **Serviços** (menu à esquerda).
2. Procure **Drive API**, selecione e clique em **Adicionar**.
3. Pronto — não precisa configurar mais nada, isso libera o `Drive.Files.insert(...)` usado no `Code.gs` para converter o PDF em texto.

## Passo 4 — Publicar como Web App
1. No editor do Apps Script, clique em **Implantar** → **Nova implantação**.
2. Em "Tipo", escolha **App da Web**.
3. Em "Executar como": **Eu (sua conta)**.
4. Em "Quem pode acessar": escolha conforme sua necessidade —
   - **Qualquer pessoa com o link** → motoristas de fora da organização também acessam sem login Google.
   - **Qualquer pessoa na [sua organização]** → exige estar logado com e-mail da empresa (mais seguro).
5. Clique em **Implantar**, autorize as permissões (vai pedir acesso ao Drive e Planilhas — é esperado, é o que faz o app funcionar).
6. Copie o **link do App da Web** gerado. Esse é o link que você vai usar/compartilhar (pode salvar como atalho na tela inicial do celular).

## Passo 5 — Testar
1. Abra o link, selecione um motorista de teste, tire fotos, preencha o checklist, assine e envie.
2. Confira se a pasta `Vistoria dd-mm-aaaa_HH-mm` apareceu dentro de `Fotos Vistoria` na pasta do motorista.
3. Uma planilha chamada **"Registro de Vistorias"** é criada automaticamente no seu Drive na primeira vez — é o seu histórico central de todas as vistorias.

## Atualizando depois de publicado
Sempre que editar `Code.gs` ou `Index.html`, é preciso ir em **Implantar → Gerenciar implantações → editar (ícone de lápis) → Nova versão → Implantar** para o link já publicado refletir as mudanças.

## Preenchimento automático a partir do contrato
- Quando o campo "Nome do cliente/motorista" perde o foco (ou ao clicar em "Buscar dados do contrato"), o app procura na pasta do cliente um arquivo cujo nome contenha **"Contrato"**, faz OCR do PDF e tenta extrair: telefone, endereço, veículo, placa, cor, ano, chassi e número do motor.
- Os campos são preenchidos automaticamente, mas continuam **editáveis** — sempre confira antes de enviar, principalmente na primeira vez.
- **Os padrões de busca (regex) foram montados com base nos rótulos do anexo de vistoria** (`Cliente:`, `Tel:`, `Endereço:`, `Placa:`, `Veículo:`, `Cor:`, `Ano:`, `Chassi:`, `Motor:`). Se o texto do contrato usar outros rótulos, o app não vai achar o valor certo. Nesse caso, me envie um exemplo do contrato (pode remover dados sensíveis) que eu ajusto a constante `CONTRACT_FIELD_PATTERNS` no `Code.gs` para bater exatamente com o layout de vocês.
- O resultado do OCR fica em cache por 6 horas (por arquivo), então buscas repetidas para o mesmo contrato não recarregam toda vez — só refaz o OCR se o cache expirar.
- Se o app não achar a pasta do cliente ou nenhum arquivo com "Contrato" no nome, ele avisa na tela e os campos continuam em branco pra preenchimento manual, sem travar o formulário.

## Personalizações rápidas
- **Cores da empresa**: no topo do `Index.html`, dentro de `:root`, troque `--cor-primaria` pela cor da sua marca (ex: o azul da ATIVUZ).
- **Logo**: troque o bloco `<div class="logo-placeholder">LOGO</div>` por `<img src="URL_PUBLICA_DO_SEU_LOGO" style="height:34px">`.
- **Itens de acessórios**: edite a lista `ACCESSORY_ITEMS` no `Code.gs`.
- **Categorias de foto**: edite a lista `PHOTO_CATEGORIES` no `Code.gs` (hoje: Painel/Hodômetro, Frente, Traseira, Lateral esquerda, Lateral direita, Interior, Danos e avarias).

## O que mudou em relação ao papel
- O desenho do carro para marcar avarias à mão foi substituído por **fotos organizadas por área do veículo** — mais rápido de preencher no celular e mais fácil de auditar depois (cada foto já entra com o nome da categoria).
- As **luzes do painel acesas** não são mais marcadas item a item: a foto da categoria "Painel / Hodômetro" (tirada com o veículo ligado) já registra isso visualmente, com evidência mais confiável do que uma marcação manual.
- Toda foto recebe automaticamente uma **marca d'água com data e hora** (canto inferior direito), carimbada no momento em que a foto é selecionada no formulário — funciona tanto para fotos tiradas na hora quanto vindas da galeria.
- As duas assinaturas (cliente e responsável) são coletadas na tela, com o dedo ou mouse.
- Cada vistoria vira uma linha na planilha "Registro de Vistorias", com link direto pra pasta de fotos daquela vistoria específica.

## Sobre a marca d'água de data/hora
- A marca usa a **data/hora do próprio celular/navegador** de quem está preenchendo — não é um carimbo criptográfico à prova de fraude, é o mesmo tipo de marca que qualquer app de câmera com timestamp usa (serve como evidência visual, não como prova legal inviolável).
- Se quiser um nível mais forte de confiabilidade (por exemplo, provar que a foto não foi enviada de uma foto antiga tirada por fora do app), dá pra evoluir depois adicionando também **geolocalização** na marca, já que ela é obtida por GPS do aparelho no momento do envio — me avise se quiser essa camada extra.
