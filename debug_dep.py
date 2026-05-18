import sys
sys.path.insert(0, '.')
from dotenv import load_dotenv; load_dotenv()
from app import _supabase

sb = _supabase()
res = sb.table('frota_veiculos').select('placa, modelo, ano_modelo, dt_aquisicao, vl_aquisicao').eq('ativo', True).order('vl_aquisicao', desc=True).execute()
rows = res.data or []

print(f"{'PLACA':<10} {'MODELO':<30} {'ANO':<6} {'DT COMPRA':<12} {'VL AQUISICAO':>15}")
print("-" * 78)
total = 0
for r in rows:
    vl = float(r.get('vl_aquisicao') or 0)
    total += vl
    placa = r.get('placa') or ''
    modelo = (r.get('modelo') or '')[:29]
    ano = str(r.get('ano_modelo') or '')
    dt = str(r.get('dt_aquisicao') or 'N/D')[:10]
    print(f"{placa:<10} {modelo:<30} {ano:<6} {dt:<12} R$ {vl:>12,.2f}")

print("-" * 78)
print(f"{'TOTAL':<49} R$ {total:>12,.2f}")
print(f"{'Depreciacao anual (total/5)':<49} R$ {total/5:>12,.2f}")
print(f"{'Depreciacao mensal (total/60)':<49} R$ {total/60:>12,.2f}")
