#!/bin/bash
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin"

REPO="/Users/karol/Documents/Dashboard-Ativuz"
FILE="planilhas/CONTAS-A-RECEBER.xlsx"

cd "$REPO" || exit 0

if ! git diff --quiet "$FILE" 2>/dev/null; then
    git add "$FILE"
    git commit -m "chore: auto-push CONTAS-A-RECEBER.xlsx"
    git push origin main
fi
