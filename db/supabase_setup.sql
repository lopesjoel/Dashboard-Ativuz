-- Execute no SQL Editor do Supabase (https://supabase.com/dashboard/project/btcwgwrajgsqndnpjman/sql)

create table if not exists public.vistorias (
  id              uuid primary key default gen_random_uuid(),
  cliente         text,
  telefone        text,
  endereco        text,
  preenchido_por  text,
  veiculo         text,
  placa           text,
  cor             text,
  ano             text,
  chassi          text,
  numero_motor    text,
  data_hora       text,
  hodometro_entrega text,
  hodometro_retorno text,
  combustivel     text,
  obs_gerais      text,
  desc_sintomas   text,
  criado_em       timestamptz default now(),
  arquivo_path    text
);

-- RLS desativado (acesso via service role key)
alter table public.vistorias disable row level security;


-- ── Contratos de Locação ──────────────────────────────────────────────────────
create table if not exists public.contratos_locacao (
  id                  uuid primary key default gen_random_uuid(),
  locatario_nome      text,
  locatario_rg        text,
  locatario_cpf       text,
  locatario_endereco  text,
  locatario_cep       text,
  locatario_telefone  text,
  avalista_nome       text,
  avalista_cpf        text,
  avalista_endereco   text,
  avalista_telefone   text,
  veiculo_descricao   text,
  veiculo_marca       text,
  veiculo_modelo      text,
  veiculo_ano         text,
  veiculo_motor       text,
  veiculo_chassi      text,
  veiculo_cor         text,
  veiculo_placa       text,
  contrato_inicio     text,
  contrato_duracao    text,
  valor_semanal       text,
  data_dia            text,
  data_mes            text,
  data_ano            text,
  testemunha1_nome    text,
  testemunha1_rg      text,
  testemunha1_cpf     text,
  testemunha2_nome    text,
  testemunha2_rg      text,
  testemunha2_cpf     text,
  arquivo_path        text,
  criado_em           timestamptz default now()
);

alter table public.contratos_locacao disable row level security;


-- ── Histórico de Documentos ───────────────────────────────────────────────────
create table if not exists public.historico_docs (
  id             uuid primary key default gen_random_uuid(),
  locatario_nome text,
  template       text,
  arquivo        text,
  data_hora      text,
  deletado       boolean default false,
  criado_em      timestamptz default now()
);

alter table public.historico_docs disable row level security;


-- ── Colunas deletado (soft delete) nas tabelas existentes ────────────────────
alter table public.contratos_locacao add column if not exists deletado boolean default false;
alter table public.vistorias         add column if not exists deletado boolean default false;
alter table public.vistorias         add column if not exists acessorios jsonb;


-- ── Usuários do sistema ───────────────────────────────────────────────────────
create table if not exists public.usuarios (
  id          uuid primary key default gen_random_uuid(),
  nome        text unique not null,
  senha_hash  text not null,
  ativo       boolean default true,
  criado_em   timestamptz default now()
);

alter table public.usuarios disable row level security;


-- ── Checklist de Contratos ────────────────────────────────────────────────────
create table if not exists public.checklist_contratos (
  id         uuid primary key default gen_random_uuid(),
  contrato   text not null,
  placa      text,
  cliente    text,
  unidade    text,
  created_at timestamptz default now()
);

alter table public.checklist_contratos disable row level security;

create table if not exists public.checklist_itens (
  id           uuid primary key default gen_random_uuid(),
  contrato_id  uuid not null references public.checklist_contratos(id) on delete cascade,
  nome         text not null,
  marcado      boolean default false,
  created_at   timestamptz default now()
);

alter table public.checklist_itens disable row level security;


-- ── Financiamentos & Consórcios ───────────────────────────────────────────────
create table if not exists public.financiamentos_contratos (
  id             uuid primary key default gen_random_uuid(),
  operacao       text not null,
  contrato       text,
  placa          text,
  data_vencimento date,
  parcelas_total  integer not null,
  valor_parcela   numeric(12,2) not null,
  created_at      timestamptz default now()
);

alter table public.financiamentos_contratos disable row level security;

create index if not exists idx_financiamentos_placa on public.financiamentos_contratos(placa);
create index if not exists idx_financiamentos_data  on public.financiamentos_contratos(data_vencimento);

-- Dados iniciais (executar uma única vez)
insert into public.financiamentos_contratos (operacao, contrato, placa, data_vencimento, parcelas_total, valor_parcela) values
('QUITADO',     'QUITADO',    'QXL-1E71',            null,         1,  33650.00),
('QUITADO',     'QUITADO',    'QMD-1C93',            null,         1,  32850.00),
('QUITADO',     'QUITADO',    'QGN-7029',            null,         1,  30728.61),
('Lucas 06',    '51434/517',  'RGE-5D71',            '2026-03-15', 1,  52426.24),
('QUITADO',     '51434/450',  'QGN-9206',            null,         1,  60778.00),
('Joel 03',     '51440/420',  'QGX-7353',            '2026-06-17', 38,  1167.42),
('Joel 09',     '51440/140',  'QGS-6G16',            '2026-06-20', 30,  1481.82),
('Lucas 09',    '51440/592',  'QGN-9356',            '2026-06-20', 30,  1481.82),
('Lucas 07',    '91362/128',  'QGU-1H54',            '2026-12-15', 50,  1171.05),
('Joel 06',     '51441/319',  '',                    '2026-10-15', 36,  1272.85),
('Lucas 05',    '51443/10',   'QGS-5J76',            '2027-01-15', 50,  1255.34),
('Andrier 01',  '9645/416',   'QGT-6I05',            '2026-12-20', 46,   897.50),
('Lucas 08',    '51441/550',  'QSE-8E63',            '2026-10-20', 36,  1305.86),
('Joel 05',     '51443/401',  'QGT-9D86',            '2027-01-20', 44,   942.81),
('Lucas 04',    '91362/217',  'QGT-1D67',            '2027-01-15', 43,   934.12),
('Joel 04',     '51443/499',  'RGE-7G92',            '2027-01-20', 44,  1071.36),
('Joel 02',     '9651/113',   'QGW-2B89',            '2027-04-27', 50,   839.74),
('Andrier 02',  '9651/825',   'RGE-7H03',            '2027-04-20', 50,  1168.39),
('Lucas 02',    '9651/879',   'QGT-9D76',            '2027-04-20', 50,   821.04),
('Ativuz 01',   '51443/594',  'RGE-3I66',            '2027-01-20', 33,  1384.89),
('Ativuz 02',   '51443/157',  'RMK-2H75',            '2027-01-20', 33,  1342.49),
('Ativuz 03',   '51443/313',  'FJV-6E53',            '2027-01-20', 33,  2093.99),
('Ativuz 04',   '91388/503',  'EZY-6303',            '2028-01-20', 42,   888.49),
('Ativuz 05',   '91388/807',  'END-0I15',            '2028-01-20', 42,   879.97),
('Ativuz BNB',  '183/002',    'EGX-2E31',            '2028-03-15', 36,  1540.24),
('Ativuz BNB',  '183/004',    'RQJ-7H29',            '2029-04-15', 48,   417.91),
('AZ BNB',      '035/003',    'ECM-1C93',            '2028-03-15', 30,  1873.79),
('Sicredi 01',  '10585560',   'EJZ-3D41 / EVE-7G53', '2029-05-12', 46,   703.51),
('AZ BNB',      '035/004',    'EWJ-2I45',            '2028-04-15', 30,  1935.83),
('AZ BNB',      '035/005',    'EXF-1F14',            '2028-04-15', 30,  1783.91),
('Ativuz 06',   '9675/395',   '',                    '2029-05-12', 44,  1418.40),
('Ativuz 07',   '9675/292',   '',                    '2029-05-12', 44,  1418.40),
('Ativuz BNB',  '183/003',    'RQI-7A69 / RQI-7A89', '2029-09-15', 48,  1001.15),
('AZ BNB',      '035/006',    'RNC-4J20',            '2027-12-15', 23,  2280.13),
('AZ BNB',      '035/007',    'ELY-4D83',            '2028-08-15', 31,  1962.57),
('Ativuz BNB',  '183/005',    '',                    '2029-02-15', 36,  1675.22),
('AGN 02',      '62970',      '',                    '2029-02-15', 36,  3678.57),
('AGN 01',      '62329',      '',                    '2029-12-15', 46,  2564.67)
on conflict do nothing;


-- ── Snapshots semanais de inadimplência ──────────────────────────────────────
create table if not exists public.inad_snapshots (
  id           uuid primary key default gen_random_uuid(),
  semana       date not null unique,   -- segunda-feira da semana
  total_casos  integer not null,
  total_valor  numeric(12,2) not null,
  criticos     integer not null default 0,
  por_etapa    jsonb,
  criado_em    timestamptz default now()
);

alter table public.inad_snapshots disable row level security;
create index if not exists idx_inad_snapshots_semana on public.inad_snapshots(semana);
