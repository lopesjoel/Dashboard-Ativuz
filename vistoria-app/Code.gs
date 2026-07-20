/*******************************************************
 * VISTORIA DE VEÍCULOS - Backend (Google Apps Script)
 * Adaptado ao modelo "Anexo I - Ordem de Serviço e Vistoria"
 * Integrado ao Dashboard Ativuz (Flask + Supabase).
 *
 * CONFIGURAÇÃO OBRIGATÓRIA: preencha as constantes abaixo
 * antes de publicar.
 *******************************************************/

// ID da pasta "2. Contratos" no Drive. Dentro dela existe uma subpasta por
// ano (ex: "! 2025", "! 2026"), e dentro de cada ano uma pasta por contrato,
// no padrão "NN/ANO_PLACA_Nome" (ex: "12/2026_QGT6I05_Matheus Sousa"). O
// script localiza a pasta certa buscando pela PLACA em todas as pastas de
// ano — não precisa saber o ano nem o "NN" de antemão.
const PARENT_FOLDER_ID = '1srRiEPkZOGaVQoF3MQcMSy_RAkGdXkfv';

// Nome da subpasta dentro da pasta do cliente onde os arquivos da
// vistoria serão guardados. Se não existir, o script cria.
const SUBFOLDER_NAME = 'Fotos Vistoria';

// (Opcional) ID de uma planilha já existente para usar como log.
// Deixe em branco ('') que o script cria uma automaticamente.
const SHEET_ID = '';

// URL base do Dashboard Ativuz (sem barra no final) e o token de API
// (mesmo valor da variável de ambiente VISTORIA_API_TOKEN no Dashboard).
// Toda vistoria enviada aqui também é gravada no Dashboard/Supabase.
const DASHBOARD_API_URL = 'https://dashboard-ativuz-two.vercel.app';
const DASHBOARD_API_TOKEN = 'j9JpsWA-4mWZbx7OcPXrGvllo65rs50ffl4qEjDjjFQ';

// Itens de acessórios/equipamentos do modelo em papel (coluna S/N/A).
// IMPORTANTE: a ordem aqui precisa bater com _CHAVES_ACC no app.py do
// Dashboard — cada posição desta lista vira o item correspondente na
// mesma posição de lá. Não reordene sem também ajustar o app.py.
const ACCESSORY_ITEMS = [
  'Calotas', 'Buzina', 'DOC. CRLV', 'Triângulo de sinalização', 'Antena', 'Sensor de ré',
  'Som/Alto-falante', 'Tapetes', 'Limpadores', 'Chave de roda', 'Vidros elétricos', 'Óleo do motor',
  'Alarme/Travas', 'Lâmpadas', 'Macaco mecânico', 'Estepe', 'Funcionamento GNV', 'Água',
  'Borracha PSG Dianteira', 'Borracha MTR Dianteira', 'Asa Urubu DD', 'Asa Urubu TD', 'Tapete de mala', 'Tampa paraquedas',
  'Borracha PSG Traseira', 'Borracha MTR Traseira', 'Asa Urubu DE', 'Asa Urubu TE', 'Bagagito', 'Lingueta'
];

// Categorias de fotos. O campo `key` precisa ser exatamente um dos ângulos
// reconhecidos pelo gerador de .docx do Dashboard (ANGULOS_FOTO em
// services/gerar_vistoria_entrada_saida.py) — não troque as chaves, só o
// `label` (texto exibido) se quiser.
const PHOTO_CATEGORIES = [
  { key: 'frontal',     label: 'Frontal' },
  { key: 'traseira',    label: 'Traseira' },
  { key: 'lateral_dir', label: 'Lateral direita' },
  { key: 'lateral_esq', label: 'Lateral esquerda' },
  { key: 'painel',      label: 'Painel / Interior' },
  { key: 'hodometro',   label: 'Hodômetro (com o veículo ligado, para registrar luzes acesas)' },
  { key: 'estepe',      label: 'Estepe' },
  { key: 'teto',        label: 'Teto / Vidros' },
  { key: 'motor',       label: 'Motor' },
  { key: 'mala',        label: 'Porta-malas' },
  { key: 'dano_1',      label: 'Dano / avaria 1 (se houver)' },
  { key: 'dano_2',      label: 'Dano / avaria 2 (se houver)' }
];

const FUEL_LEVELS = ['Vazio', '1/4', '1/2', '3/4', 'Cheio'];

/*******************************************************
 * Não é necessário editar nada abaixo desta linha.
 *******************************************************/

// Aceita ?etapa=saida&contratoId=...&vistoriaId=... na URL (usado pelo botão
// "Devolução" do Histórico de Vistorias do Dashboard, pra abrir o app já na
// etapa de saída, com o contrato certo travado, vinculado à vistoria certa).
function doGet(e) {
  const params = (e && e.parameter) || {};
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.presetEtapa = params.etapa === 'saida' ? 'saida' : '';
  tpl.presetContratoId = params.contratoId || '';
  tpl.presetVistoriaId = params.vistoriaId || '';
  return tpl.evaluate()
    .setTitle('Ordem de Serviço e Vistoria')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Retorna toda a configuração que o front-end precisa montar o formulário
function getFormConfig() {
  return {
    accessoryItems: ACCESSORY_ITEMS,
    photoCategories: PHOTO_CATEGORIES,
    fuelLevels: FUEL_LEVELS,
    contratos: getContratosAtivos_()
  };
}

// Busca a lista de contratos ativos no Dashboard (para vincular a vistoria
// a um contrato real). Se o Dashboard estiver fora do ar ou mal configurado,
// retorna lista vazia — o formulário continua funcionando sem vínculo de
// contrato (preenchimento manual do nome do cliente).
function getContratosAtivos_() {
  if (!DASHBOARD_API_URL || DASHBOARD_API_URL.indexOf('COLOQUE_AQUI') === 0) return [];
  try {
    const resp = UrlFetchApp.fetch(DASHBOARD_API_URL + '/api/contratos/ativos', {
      method: 'get',
      headers: { 'X-Vistoria-Token': DASHBOARD_API_TOKEN },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return [];
    const data = JSON.parse(resp.getContentText());
    return data.contratos || [];
  } catch (err) {
    Logger.log('Erro ao buscar contratos ativos: ' + err.message);
    return [];
  }
}

// Remove tudo que não for letra/número e deixa maiúsculo, pra comparar placas
// sem depender de hífen/espaço (ex: "ECZ-8C02" e "ecz8c02" batem).
function normalizePlaca_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// Procura, em todas as subpastas de ano (ex: "! 2025") dentro de
// PARENT_FOLDER_ID, uma pasta de contrato cujo nome contenha a placa
// informada. Não apaga nem move nada — só lê a árvore de pastas.
function findContractFolderByPlaca_(placa) {
  const normPlaca = normalizePlaca_(placa);
  if (!normPlaca) return null;

  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const yearFolders = [];
  const it = parent.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    const m = f.getName().trim().match(/^!\s*(\d{4})$/);
    if (m) yearFolders.push({ folder: f, year: parseInt(m[1], 10) });
  }
  // Anos mais recentes primeiro (contratos ativos tendem a ser recentes)
  yearFolders.sort(function (a, b) { return b.year - a.year; });

  for (let i = 0; i < yearFolders.length; i++) {
    const subs = yearFolders[i].folder.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      if (normalizePlaca_(sub.getName()).indexOf(normPlaca) !== -1) {
        return sub;
      }
    }
  }
  return null;
}

// Localiza a pasta do contrato pela placa. Se não achar (ex: contrato novo,
// pasta ainda não criada por quem organiza o Drive), cria uma pasta nova
// dentro do ano atual — nunca mexe em pastas/arquivos já existentes.
function findOrCreateContractFolder_(placa, clientName) {
  const found = findContractFolderByPlaca_(placa);
  if (found) return found;

  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const yearName = '! ' + new Date().getFullYear();
  const yearFolder = getOrCreateSubfolder_(parent, yearName);
  const safeName = (placa ? placa.toUpperCase() : 'SEM-PLACA') + '_' + (clientName || 'Cliente');
  return yearFolder.createFolder(safeName);
}

function getOrCreateSubfolder_(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(name);
}

/**
 * Busca telefone/endereço/cor/ano/chassi/nº motor do contrato pela placa,
 * direto da planilha do Dashboard (/api/contrato/dados). Chamada pelo
 * front-end quando o usuário seleciona um contrato da lista.
 * Retorna { found: false, reason } ou
 * { found: true, phone, address, color, year, chassis, motorNumber }
 */
function getContractData(placa) {
  if (!placa) return { found: false, reason: 'Placa vazia' };
  if (!DASHBOARD_API_URL || DASHBOARD_API_URL.indexOf('COLOQUE_AQUI') === 0) {
    return { found: false, reason: 'DASHBOARD_API_URL não configurado' };
  }
  try {
    const resp = UrlFetchApp.fetch(
      DASHBOARD_API_URL + '/api/contrato/dados?placa=' + encodeURIComponent(placa),
      {
        method: 'get',
        headers: { 'X-Vistoria-Token': DASHBOARD_API_TOKEN },
        muteHttpExceptions: true
      }
    );
    if (resp.getResponseCode() !== 200) {
      return { found: false, reason: 'HTTP ' + resp.getResponseCode() };
    }
    return JSON.parse(resp.getContentText());
  } catch (err) {
    return { found: false, reason: 'Erro ao buscar dados do contrato: ' + err.message };
  }
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
      'Data/Hora', 'Etapa', 'Contrato', 'Cliente', 'Telefone', 'Endereço', 'Preenchido por',
      'Veículo', 'Placa', 'Cor', 'Ano', 'Chassi', 'Motor',
      'Hodômetro', 'Combustível',
      'Acessórios com problema (N/A)',
      'Observações', 'Descrição dos sintomas', 'Qtde fotos', 'Link da pasta',
      'Dashboard: status'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Função principal chamada pelo front-end.
 * payload = {
 *   etapa: 'entrada' | 'saida',
 *   contratoId (opcional), vistoriaId (opcional, só na saída — identifica
 *   exatamente qual vistoria completar, em vez de depender só do contrato),
 *   clientName, phone, address, filledBy,
 *   vehicle, plate, color, year, chassis, motorNumber,
 *   odometer, fuelLevel,
 *   accessories: [{item, status}],       // status: S | N | A
 *   observations, symptoms,
 *   clientSignature: {data, mimeType},
 *   responsibleSignature: {data, mimeType},
 *   photos: [{category, name, mimeType, data}]  // category = chave de PHOTO_CATEGORIES, data em base64
 * }
 */
function uploadVistoria(payload) {
  if (!payload || !payload.contratoId) {
    throw new Error('Selecione um contrato da lista.');
  }
  if (!payload.plate) {
    throw new Error('Placa não encontrada — selecione o contrato novamente.');
  }
  if (!payload.photos || payload.photos.length === 0) {
    throw new Error('Adicione ao menos uma foto.');
  }

  const clientFolder = findOrCreateContractFolder_(payload.plate, payload.clientName);
  const photosRootFolder = getOrCreateSubfolder_(clientFolder, SUBFOLDER_NAME);

  const now = new Date();
  const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy_HH-mm');
  const etapaLabel = payload.etapa === 'saida' ? 'Saida' : 'Entrada';
  const inspectionFolder = photosRootFolder.createFolder('Vistoria ' + etapaLabel + ' ' + stamp);

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

  // Envia também pro Dashboard (Supabase) — gera o .docx oficial e aparece
  // no Histórico de Vistorias. Se falhar, a vistoria continua salva no
  // Drive/planilha normalmente; o erro só é reportado ao usuário.
  const dashboardResult = sendToDashboard_(payload);

  const sheet = getOrCreateLogSheet_();
  sheet.appendRow([
    now,
    etapaLabel,
    payload.contratoId || '',
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
    payload.odometer || '',
    payload.fuelLevel || '',
    problemAccessories || 'Nenhum',
    payload.observations || '',
    payload.symptoms || '',
    payload.photos.length,
    inspectionFolder.getUrl(),
    dashboardResult.ok ? 'OK' : ('Falhou: ' + dashboardResult.error)
  ]);

  return {
    ok: true,
    folderUrl: inspectionFolder.getUrl(),
    dashboard: dashboardResult
  };
}

// Envia a vistoria pro Dashboard Ativuz (Flask + Supabase). Não interrompe
// o fluxo principal em caso de falha — só reporta o resultado.
function sendToDashboard_(payload) {
  if (!DASHBOARD_API_URL || DASHBOARD_API_URL.indexOf('COLOQUE_AQUI') === 0) {
    return { ok: false, error: 'DASHBOARD_API_URL não configurado' };
  }
  try {
    const resp = UrlFetchApp.fetch(DASHBOARD_API_URL + '/api/vistoria/importar', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Vistoria-Token': DASHBOARD_API_TOKEN },
      payload: JSON.stringify({
        etapa: payload.etapa === 'saida' ? 'saida' : 'entrada',
        contratoId: payload.contratoId || '',
        vistoriaId: payload.vistoriaId || '',
        clientName: payload.clientName,
        phone: payload.phone,
        address: payload.address,
        filledBy: payload.filledBy,
        vehicle: payload.vehicle,
        plate: payload.plate,
        color: payload.color,
        year: payload.year,
        chassis: payload.chassis,
        motorNumber: payload.motorNumber,
        odometer: payload.odometer,
        fuelLevel: payload.fuelLevel,
        observations: payload.observations,
        symptoms: payload.symptoms,
        accessories: payload.accessories,
        photos: payload.photos.map(function (p) {
          return { category: p.category, mimeType: p.mimeType, data: p.data };
        }),
        clientSignature: payload.clientSignature,
        responsibleSignature: payload.responsibleSignature
      }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      return { ok: true, response: JSON.parse(resp.getContentText()) };
    }
    return { ok: false, error: 'HTTP ' + code + ': ' + resp.getContentText() };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
