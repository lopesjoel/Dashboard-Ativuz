import sys
sys.path.insert(0, '.')
from dotenv import load_dotenv; load_dotenv()
from app import _supabase

sb = _supabase()

res = sb.table("carteira_judicializada").select("*").limit(1).execute()
if not res.data:
    print("Sem registros."); exit()

r = res.data[0]
status_original = r['status']
status_novo = 'Acordo' if status_original != 'Acordo' else 'Em Análise'

print(f"Status original: {status_original!r}")
print(f"Tentando mudar para: {status_novo!r}")

upd = sb.table("carteira_judicializada").update({
    "cliente":       r["cliente"],
    "cpf_cnpj":      r["cpf_cnpj"] or "",
    "inicio_divida": r["inicio_divida"],
    "valor_atual":   r["valor_atual"],
    "status":        status_novo,
    "num_processo":  r["num_processo"] or "",
    "atualizado_em": "now()",
}).eq("id", r["id"]).execute()

print(f"Retorno do update: {upd.data}")

# Busca novamente para confirmar persistência
res2 = sb.table("carteira_judicializada").select("status").eq("id", r["id"]).execute()
print(f"Status após update: {res2.data[0]['status']!r}")

# Reverte
sb.table("carteira_judicializada").update({"status": status_original}).eq("id", r["id"]).execute()
print(f"Revertido para: {status_original!r}")
