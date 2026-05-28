@"
# Agent skills portabili

Questa cartella contiene copie versionate e portabili delle skill operative usate nel progetto Fantozzi.

La cartella `.claude/skills/` può restare come copia locale di lavoro, ma è ignorata da Git.
Quando una skill viene modificata in `.claude/skills/`, la versione aggiornata deve essere copiata qui prima del commit.

Regole:
- non inserire credenziali, password, token o API key;
- non inserire dati reali, XML, estratti bancari, comprovantes o output;
- non committare file da `output/`, `Database_Dev/`, `bank_raw/` o `reports/`;
- le skill devono essere autosufficienti.
"@ | Set-Content agent-skills/README.md -Encoding UTF8