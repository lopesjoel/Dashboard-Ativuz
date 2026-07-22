-- Adiciona distinção Consórcio x Financiamento e registro de resgate de
-- consórcio em financiamentos_contratos. Execute no Supabase SQL Editor.

alter table public.financiamentos_contratos
  add column if not exists tipo text not null default 'financiamento'
    check (tipo in ('financiamento', 'consorcio')),
  add column if not exists valor_resgate numeric(12,2),
  add column if not exists data_resgate  date;

-- Classificação dos registros já existentes: Sicredi, AZ BNB e AGN são
-- financiamento; todo o resto (inclusive os QUITADO e a Ativuz BNB) é
-- consórcio.
update public.financiamentos_contratos
set tipo = 'consorcio'
where not (
  operacao ilike '%sicredi%'
  or operacao = 'AZ BNB'
  or operacao ilike '%agn%'
);
