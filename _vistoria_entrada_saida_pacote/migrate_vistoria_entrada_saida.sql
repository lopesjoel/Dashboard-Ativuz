-- ─────────────────────────────────────────────────────────────────────────
-- Migration: tabela `vistorias` passa a ter ENTRADA e SAÍDA no mesmo registro
-- Executar no SQL Editor do Supabase
-- ─────────────────────────────────────────────────────────────────────────

-- 1. Adicionar colunas de ENTRADA (renomear as antigas é mais arriscado;
--    aqui criamos novas e migramos os dados existentes ao final)
alter table public.vistorias
  add column if not exists contrato_id           text,
  add column if not exists data_entrada          text,
  add column if not exists hodometro_entrada     text,
  add column if not exists combustivel_entrada   text,
  add column if not exists obs_entrada           text,
  add column if not exists sintomas_entrada      text,
  add column if not exists responsavel_entrada   text,
  add column if not exists acessorios_entrada    jsonb,
  add column if not exists fotos_entrada         text[],

-- 2. Colunas de SAÍDA (devolução)
  add column if not exists data_saida            text,
  add column if not exists hodometro_saida       text,
  add column if not exists combustivel_saida     text,
  add column if not exists obs_saida             text,
  add column if not exists sintomas_saida        text,
  add column if not exists responsavel_saida     text,
  add column if not exists acessorios_saida      jsonb,
  add column if not exists fotos_saida           text[],

-- 3. Controle do ciclo
  add column if not exists status                text default 'pendente_saida',
  --   valores possíveis: 'pendente_saida' | 'completa' | 'cancelada'
  add column if not exists divergencias          jsonb,
  --   [{"item":"Calotas","entrada":"S","saida":"N","motivo":"Item ausente na devolução"}, ...]
  add column if not exists arquivo_entrada_path  text,
  add column if not exists arquivo_completo_path text,
  add column if not exists atualizado_em         timestamptz default now();

-- 4. Constraint do status (só os valores permitidos)
do $$
begin
  if not exists (
    select 1 from pg_constraint where conname = 'vistorias_status_chk'
  ) then
    alter table public.vistorias
      add constraint vistorias_status_chk
      check (status in ('pendente_saida', 'completa', 'cancelada'));
  end if;
end$$;

-- 5. Índice para buscar por contrato/placa rapidamente
create index if not exists idx_vistorias_contrato on public.vistorias(contrato_id);
create index if not exists idx_vistorias_placa    on public.vistorias(placa);
create index if not exists idx_vistorias_status   on public.vistorias(status);

-- 6. Trigger para manter `atualizado_em` em sincronia (opcional, recomendado)
create or replace function public.tg_vistorias_set_atualizado_em()
returns trigger language plpgsql as $$
begin
  new.atualizado_em := now();
  return new;
end$$;

drop trigger if exists trg_vistorias_atualizado_em on public.vistorias;
create trigger trg_vistorias_atualizado_em
  before update on public.vistorias
  for each row execute function public.tg_vistorias_set_atualizado_em();

-- 7. Migração leve: copiar dados antigos para os novos campos de ENTRADA
update public.vistorias
   set data_entrada        = coalesce(data_entrada,        data_hora),
       hodometro_entrada   = coalesce(hodometro_entrada,   hodometro_entrega),
       hodometro_saida     = coalesce(hodometro_saida,     hodometro_retorno),
       combustivel_entrada = coalesce(combustivel_entrada, combustivel),
       obs_entrada         = coalesce(obs_entrada,         obs_gerais),
       sintomas_entrada    = coalesce(sintomas_entrada,    desc_sintomas),
       acessorios_entrada  = coalesce(acessorios_entrada,  acessorios),
       arquivo_entrada_path= coalesce(arquivo_entrada_path,arquivo_path),
       status              = case
                                when hodometro_retorno is not null and hodometro_retorno <> ''
                                  then 'completa'
                                else 'pendente_saida'
                             end
 where coalesce(data_entrada, '') = ''
    or coalesce(acessorios_entrada::text, '') = '';

-- 8. (Opcional) As colunas antigas continuam existindo para compatibilidade
--    durante a transição. Quando o app estiver 100% migrado, dá para
--    fazer um `alter table … drop column hodometro_entrega;` etc.

-- ─────────────────────────────────────────────────────────────────────────
-- LGPD — RLS mínimo (recomendado mesmo no uso interno)
-- ─────────────────────────────────────────────────────────────────────────
-- Reative o RLS quando quiser segregar acesso por usuário:
--
--   alter table public.vistorias enable row level security;
--
--   create policy vistorias_owner_select on public.vistorias
--     for select using (auth.uid() is not null);
--
--   create policy vistorias_owner_insert on public.vistorias
--     for insert with check (auth.uid() is not null);
--
--   create policy vistorias_owner_update on public.vistorias
--     for update using (auth.uid() is not null);
-- ─────────────────────────────────────────────────────────────────────────
