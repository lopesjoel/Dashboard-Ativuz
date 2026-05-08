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
