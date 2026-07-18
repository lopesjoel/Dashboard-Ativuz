/*******************************************************
 * VISTORIA DE VEÍCULOS - Backend (Google Apps Script)
 * Adaptado ao modelo "Anexo I - Ordem de Serviço e Vistoria"
 *
 * CONFIGURAÇÃO OBRIGATÓRIA: preencha as constantes abaixo
 * antes de publicar.
 *******************************************************/

// ID da pasta do Drive que contém as pastas de cada cliente/motorista.
// Pegue o ID na URL da pasta: drive.google.com/drive/folders/ESTE_TRECHO_AQUI
const PARENT_FOLDER_ID = 'COLOQUE_AQUI_O_ID_DA_PASTA_MAE';

// Nome da subpasta dentro da pasta do cliente onde os arquivos da
// vistoria serão guardados. Se não existir, o script cria.
const SUBFOLDER_NAME = 'Fotos Vistoria';

// (Opcional) ID de uma planilha já existente para usar como log.
// Deixe em branco ('') que o script cria uma automaticamente.
const SHEET_ID = '';

// Itens de acessórios/equipamentos do modelo em papel (coluna S/N/A)
const ACCESSORY_ITEMS = [
  'Calotas', 'Buzina', 'DOC. CRLV', 'Triângulo de sinalização', 'Antena', 'Sensor de ré',
  'Som/Alto-falante', 'Tapetes', 'Limpadores', 'Chave de roda', 'Vidros elétricos', 'Óleo do motor',
  'Alarme/Travas', 'Lâmpadas', 'Macaco mecânico', 'Estepe', 'Funcionamento GNV', 'Água',
  'Borracha PSG Dianteira', 'Borracha MTR Dianteira', 'Asa Urubu DD', 'Asa Urubu TD', 'Tapete de mala', 'Tampa paraquedas',
  'Borracha PSG Traseira', 'Borracha MTR Traseira', 'Asa Urubu DE', 'Asa Urubu TE', 'Bagagito', 'Lingueta'
];

// Categorias de fotos (substituem o desenho do carro para marcar avarias
// e também servem para registrar as luzes acesas no painel)
const PHOTO_CATEGORIES = [
  'Painel / Hodômetro',
  'Frente',
  'Traseira',
  'Lateral esquerda',
  'Lateral direita',
  'Interior / Bancos',
  'Danos e avarias'
];

const FUEL_LEVELS = ['Vazio', '1/4', '1/2', '3/4', 'Cheio'];

// Trecho do nome do arquivo do contrato dentro da pasta do cliente/motorista
// (busca é "contém", não precisa ser o nome exato)
const CONTRACT_FILENAME_HINT = 'Contrato';

// Padrões (regex) usados para extrair cada campo do texto do contrato via OCR.
// IMPORTANTE: estes padrões foram montados com base nos rótulos vistos no
// anexo de vistoria. Se o CONTRATO principal usar rótulos diferentes, ajuste
// aqui (pode adicionar mais de um padrão por campo — o primeiro que bater é usado).
const CONTRACT_FIELD_PATTERNS = {
  phone: [/Tel(?:efone)?:?\s*([^\n]+)/i, /Fone:?\s*([^\n]+)/i],
  address: [/Endere[cç]o:?\s*([^\n]+)/i],
  vehicle: [/Ve[ií]culo:?\s*([^\n]+)/i],
  plate: [/Placa:?\s*([^\n]+)/i],
  color: [/\bCor:?\s*([^\n]+)/i],
  year: [/\bAno:?\s*([^\n\/]+\/?\d*)/i],
  chassis: [/Chassi:?\s*([^\n]+)/i],
  motorNumber: [/Motor:?\s*([^\n]+)/i]
};

/*******************************************************
 * Não é necessário editar nada abaixo desta linha.
 *******************************************************/

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Ordem de Serviço e Vistoria')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Retorna toda a configuração que o front-end precisa montar o formulário
function getFormConfig() {
  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const folders = parent.getFolders();
  const clients = [];
  while (folders.hasNext()) {
    clients.push(folders.next().getName());
  }
  clients.sort(function (a, b) { return a.localeCompare(b, 'pt-BR'); });

  return {
    clients: clients,
    accessoryItems: ACCESSORY_ITEMS,
    photoCategories: PHOTO_CATEGORIES,
    fuelLevels: FUEL_LEVELS
  };
}

// Localiza a pasta do cliente/motorista pelo nome, ou cria se não existir
function findOrCreateClientFolder_(clientName) {
  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const existing = parent.getFoldersByName(clientName);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parent.createFolder(clientName);
}

function getOrCreateSubfolder_(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(name);
}

/**
 * Busca o contrato na pasta do cliente/motorista e extrai os dados dele.
 * Chamada pelo front-end quando o usuário escolhe/confirma o nome do cliente.
 * Retorna { found: false } se não achar pasta/arquivo, ou
 * { found: true, phone, address, vehicle, plate, color, year, chassis, motorNumber }
 */
function getContractData(clientName) {
  if (!clientName) return { found: false, reason: 'Nome vazio' };

  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const clientFolders = parent.getFoldersByName(clientName);
  if (!clientFolders.hasNext()) {
    return { found: false, reason: 'Pasta do cliente não encontrada' };
  }
  const clientFolder = clientFolders.next();

  const files = clientFolder.searchFiles(
    "title contains '" + CONTRACT_FILENAME_HINT.replace(/'/g, "\\'") + "'"
  );
  if (!files.hasNext()) {
    return { found: false, reason: 'Nenhum arquivo com "' + CONTRACT_FILENAME_HINT + '" no nome' };
  }
  const contractFile = files.next();

  try {
    const text = ocrFileToText_(contractFile.getId());
    return {
      found: true,
      phone: extractField_(text, CONTRACT_FIELD_PATTERNS.phone),
      address: extractField_(text, CONTRACT_FIELD_PATTERNS.address),
      vehicle: extractField_(text, CONTRACT_FIELD_PATTERNS.vehicle),
      plate: extractField_(text, CONTRACT_FIELD_PATTERNS.plate),
      color: extractField_(text, CONTRACT_FIELD_PATTERNS.color),
      year: extractField_(text, CONTRACT_FIELD_PATTERNS.year),
      chassis: extractField_(text, CONTRACT_FIELD_PATTERNS.chassis),
      motorNumber: extractField_(text, CONTRACT_FIELD_PATTERNS.motorNumber)
    };
  } catch (err) {
    return { found: false, reason: 'Erro ao ler o contrato: ' + err.message };
  }
}

function extractField_(text, patterns) {
  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match && match[1]) {
      return match[1].trim().replace(/\s{2,}/g, ' ');
    }
  }
  return '';
}

// Converte o PDF em texto via OCR do Google Drive (usa o Advanced Drive Service).
// Resultado fica em cache por algumas horas para não reprocessar toda hora.
function ocrFileToText_(fileId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ocr_' + fileId;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const blob = DriveApp.getFileById(fileId).getBlob();
  const resource = { title: 'OCR_temp_' + fileId, mimeType: MimeType.GOOGLE_DOCS };
  const ocrFile = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'pt' });

  const doc = DocumentApp.openById(ocrFile.id);
  const text = doc.getBody().getText();

  // limpa o arquivo temporário gerado pelo OCR
  DriveApp.getFileById(ocrFile.id).setTrashed(true);

  // cache por 6 horas (21600s) — o contrato não muda com frequência
  cache.put(cacheKey, text, 21600);

  return text;
}

function getOrCreateLogSheet_() {
  let ss;
  if (SHEET_ID) {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } else {
    const props = PropertiesService.getScriptProperties();
    const savedId = props.getProperty('LOG_SHEET_ID');
    if (savedId) {
      ss = SpreadsheetApp.openById(savedId);
    } else {
      ss = SpreadsheetApp.create('Registro de Vistorias');
      props.setProperty('LOG_SHEET_ID', ss.getId());
      Logger.log('Planilha de log criada: ' + ss.getUrl());
    }
  }
  let sheet = ss.getSheetByName('Vistorias');
  if (!sheet) {
    sheet = ss.insertSheet('Vistorias');
    sheet.appendRow([
      'Data/Hora', 'Cliente', 'Telefone', 'Endereço', 'Preenchido por',
      'Veículo', 'Placa', 'Cor', 'Ano', 'Chassi', 'Motor',
      'Hodômetro entrega', 'Hodômetro retorno', 'Combustível',
      'Acessórios com problema (N/A)',
      'Observações', 'Descrição dos sintomas', 'Qtde fotos', 'Link da pasta'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Função principal chamada pelo front-end.
 * payload = {
 *   clientName, phone, address, filledBy,
 *   vehicle, plate, color, year, chassis, motorNumber,
 *   odometerDelivery, odometerReturn, fuelLevel,
 *   accessories: [{item, status}],       // status: S | N | A
 *   observations, symptoms,
 *   clientSignature: {data, mimeType},
 *   responsibleSignature: {data, mimeType},
 *   photos: [{category, name, mimeType, data}]  // data em base64
 * }
 */
function uploadVistoria(payload) {
  if (!payload || !payload.clientName) {
    throw new Error('Informe o cliente/motorista.');
  }
  if (!payload.photos || payload.photos.length === 0) {
    throw new Error('Adicione ao menos uma foto.');
  }

  const clientFolder = findOrCreateClientFolder_(payload.clientName);
  const photosRootFolder = getOrCreateSubfolder_(clientFolder, SUBFOLDER_NAME);

  const now = new Date();
  const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy_HH-mm');
  const inspectionFolder = photosRootFolder.createFolder('Vistoria ' + stamp);

  // Salva as fotos, prefixando com a categoria
  payload.photos.forEach(function (photo, i) {
    const safeCategory = (photo.category || 'Foto').replace(/[\/\\]/g, '-');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(photo.data),
      photo.mimeType,
      (i + 1) + ' - ' + safeCategory + ' - ' + photo.name
    );
    inspectionFolder.createFile(blob);
  });

  // Salva as assinaturas
  if (payload.clientSignature && payload.clientSignature.data) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(payload.clientSignature.data),
      payload.clientSignature.mimeType || 'image/png',
      'assinatura_cliente.png'
    );
    inspectionFolder.createFile(blob);
  }
  if (payload.responsibleSignature && payload.responsibleSignature.data) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(payload.responsibleSignature.data),
      payload.responsibleSignature.mimeType || 'image/png',
      'assinatura_responsavel.png'
    );
    inspectionFolder.createFile(blob);
  }

  // Lista de acessórios marcados como N (não existente) ou A (avariado)
  const problemAccessories = (payload.accessories || [])
    .filter(function (a) { return a.status === 'N' || a.status === 'A'; })
    .map(function (a) { return a.item + ' (' + a.status + ')'; })
    .join('; ');

  const sheet = getOrCreateLogSheet_();
  sheet.appendRow([
    now,
    payload.clientName,
    payload.phone || '',
    payload.address || '',
    payload.filledBy || '',
    payload.vehicle || '',
    payload.plate || '',
    payload.color || '',
    payload.year || '',
    payload.chassis || '',
    payload.motorNumber || '',
    payload.odometerDelivery || '',
    payload.odometerReturn || '',
    payload.fuelLevel || '',
    problemAccessories || 'Nenhum',
    payload.observations || '',
    payload.symptoms || '',
    payload.photos.length,
    inspectionFolder.getUrl()
  ]);

  return {
    ok: true,
    folderUrl: inspectionFolder.getUrl()
  };
}
