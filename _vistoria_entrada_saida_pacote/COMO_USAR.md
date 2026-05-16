# Vistoria Entrada × Saída — Guia de uso

Pacote com 3 arquivos para integrar ao projeto Ativuz:

```
docx_templates/VISTORIA_ENTRADA_SAIDA_TEMPLATE.docx   ← template novo
gerar_vistoria_entrada_saida.py                       ← gerador (raiz do projeto)
migrate_vistoria_entrada_saida.sql                    ← migration do Supabase
```

E dois arquivos de apoio:

```
criar_template.py          ← gera o template (rodar uma vez; já incluído o resultado)
VISTORIA_EXEMPLO.docx      ← exemplo gerado (entrega + devolução com 4 divergências)
VISTORIA_EXEMPLO.pdf       ← mesmo exemplo em PDF
VISTORIA_PARCIAL.docx      ← exemplo só com a entrega (status pendente_saida)
```

---

## 1. Aplicar a migration no Supabase

Cole `migrate_vistoria_entrada_saida.sql` no SQL Editor e execute. Ele:

- Adiciona colunas `*_entrada` e `*_saida` à tabela `vistorias`.
- Adiciona `status` (`pendente_saida` | `completa` | `cancelada`), `divergencias` (jsonb) e
  paths separados (`arquivo_entrada_path`, `arquivo_completo_path`).
- Copia os dados antigos (`hodometro_entrega` → `hodometro_entrada`, etc.) para os novos
  campos, sem perder nada.
- Cria índices em `contrato_id`, `placa`, `status`.

> As colunas antigas continuam existindo. Quando o app estiver 100% migrado, dá para
> dropar (`alter table … drop column hodometro_entrega`).

## 2. Colocar o template e o script na pasta certa

```
projeto_ativuz/
├── gerar_vistoria_entrada_saida.py             ← copiar para a raiz, ao lado de gerar_contrato.py
└── docx_templates/
    └── VISTORIA_ENTRADA_SAIDA_TEMPLATE.docx   ← copiar para cá
```

## 3. Como chamar no app.py

```python
from gerar_vistoria_entrada_saida import gerar_vistoria_entrada_saida, docx_para_pdf

# Quando o cliente RETIRA o carro (vistoria de entrada):
dados = {
    "contrato_id": "0042",
    "cliente_nome": request.form.get("cliente_nome"),
    # ... demais campos do cliente/veículo ...
    "data_entrada": agora.strftime("%d/%m/%Y %H:%M"),
    "hodometro_entrada": request.form.get("hodometro"),
    "combustivel_entrada": request.form.get("combustivel"),
    "obs_entrada": request.form.get("observacoes"),
    "responsavel_entrada": session.get("usuario"),
    "acessorios_entrada": {
        f"acc_{k}": request.form.get(f"acc_{k}", "") for k in [
            "calotas","buzina","doc_crlv", ...  # lista padrão
        ]
    },
    "fotos_entrada": fotos_paths,
}
resumo = gerar_vistoria_entrada_saida(
    dados,
    caminho_saida=f"contratos/VISTORIA_{placa}_{contrato_id}.docx",
)
# Salva no Supabase com status='pendente_saida'
sb.table("vistorias").insert({
    **dados, "status": resumo["status"],
    "arquivo_entrada_path": resumo["arquivo"],
}).execute()

# Quando o cliente DEVOLVE o carro:
# 1) Carregue o registro existente: sb.table("vistorias").select("*").eq("contrato_id", X)
# 2) Acrescente os campos de SAÍDA ao dict `dados`
# 3) Chame `gerar_vistoria_entrada_saida(dados, mesmo_arquivo)` — o mesmo PDF é
#    regenerado já com as duas colunas preenchidas e com as divergências em vermelho.
# 4) Atualize no banco:
sb.table("vistorias").update({
    "data_saida": dados["data_saida"],
    # ... resto dos _saida ...
    "acessorios_saida": dados["acessorios_saida"],
    "fotos_saida": dados["fotos_saida"],
    "status": resumo["status"],         # vira "completa"
    "divergencias": resumo["divergencias"],
    "arquivo_completo_path": resumo["arquivo"],
}).eq("contrato_id", contrato_id).execute()
```

## 4. O que o script retorna

```python
{
    "arquivo":      "contratos/VISTORIA_X.docx",
    "status":       "pendente_saida" | "completa",
    "divergencias": [
        ("Calotas",    "S", "N", "Item ausente na devolução"),
        ("Tapetes",    "S", "A", "Avariado na devolução"),
        ...
    ],
}
```

`divergencias` é a lista de itens que mudaram entre a entrega e a devolução — útil para
mostrar um resumo na tela e fundamentar cobrança de danos.

## 5. Sugestão de UX no front-end

- Tela `/vistoria/<contrato_id>`: mostra status atual e abre o formulário do lado
  correto (entrega ou devolução). Os campos do lado já preenchido ficam read-only.
- Quando status == "completa", a tela mostra um destaque vermelho com a lista de
  divergências e oferece um botão "Gerar termo de cobrança" (próxima evolução).
- O documento final (status="completa") é o que vale juridicamente — gere o PDF
  só nessa hora e mande para o cliente assinar (ZapSign/Clicksign).

## 6. Próximas evoluções sugeridas

1. **Galeria pareada de fotos**: hoje as fotos da entrega e da devolução aparecem em
   blocos separados. O ideal é parear por ângulo (frente entrega × frente devolução)
   e marcar visualmente as diferenças. Daria para usar a tag de cada foto (`frente`,
   `traseira`, `lateral_dir`, etc.) e renderizar uma grade 2 colunas.
2. **Assinatura digital**: integrar Clicksign / ZapSign no momento de finalizar a
   vistoria (substituir as linhas em branco por um link "Assinar pelo celular").
3. **Hash de integridade**: incluir SHA-256 do PDF no rodapé + QR Code apontando para
   uma página de verificação. Resolve disputa sobre adulteração.
4. **Comparação automática de hodômetro**: alertar se `hodometro_saida - hodometro_entrada`
   ultrapassa a franquia contratada (e calcular a cobrança).
