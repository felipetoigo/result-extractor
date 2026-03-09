Place model.xlsx here.

When you run "Import and Convert", the app will:
1. Create exported_<timestamp>.xlsx on your Desktop (full data).
2. Copy model.xlsx to copied_model_<timestamp>.xlsx on your Desktop and fill it with the table section only, mapping columns:
   ESPÉCIE → ESPECIFICAÇÃO, VENCIMENTO → VENCIMENTO, VALOR → VALOR,
   CORREÇÃO MONETÁRIA → CORREÇÃO MONETÁRIA, JUROS → JUROS, MULTA → MULTA,
   HONORÁRIOS → HONORÁRIOS ADVOCATÍCIOS, TOTAL → TOTAL.

The model sheet must have a header row containing these column names (e.g. ESPECIFICAÇÃO, VENCIMENTO, etc.).
