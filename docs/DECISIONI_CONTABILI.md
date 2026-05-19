# DECISIONI_CONTABILI.md — Programma Registrazioni Contabili / Fantozzi

Aggiornato: 2026-05-17

## Scopo

Questo file raccoglie le decisioni contabili e gestionali già consolidate nel progetto.

Il progetto non mira a sostituire la contabilità fiscale ufficiale.
Serve a generare un DRE gerencial attendibile per gestione dell’impresa, con audit trail e separazione dei casi dubbi.

## Principio generale

Le registrazioni prodotte sono gerenciais/provvisorie, non fiscali definitive.

Obiettivo:
- classificare costi/receitas/ajustes in modo utile alla gestione;
- separare chiaramente ciò che è OK da ciò che richiede revisione;
- non contaminare DRE con casi non decisi;
- mantenere tracciabilità delle fonti.

## REVISAR vs OK

Regola:
solo lançamentos con status OK e senza alert rilevanti entrano nei report OK/DRE, salvo eccezioni esplicite e testate.

I casi REVISAR devono restare fuori da:
- DRE;
- relatórios gerenciais OK;
- pré-livro OK.

## Bonificação recebida

Decisione:
bonificação recebida è riduzione di costo di acquisto.

Trattamento:
- entra nel DRE come riduzione di costo;
- conto tecnico/gerencial: `2.05.003 - Reducao de Custos - Bonificacoes Recebidas`;
- trattamento: `REDUCAO_CUSTO`.

Condizione:
deve essere riconosciuta come `BONIFICACAO_RECEBIDA` coerente.

Non fare:
- non trattarla come receita generica;
- non trattarla come costo normale;
- non lasciarla contaminare la tabella delle pendenze operative se già decisa.

## COMODATO

Decisione:
COMODATO non entra in DRE.

Trattamento:
- resta fuori dai report OK/DRE;
- può essere documentato in allegato, per esempio “Bens de terceiros em comodato”;
- non deve essere promosso come costo;
- non deve essere promosso come investimento proprio senza decisione specifica.

## CFOP 1949 / recusa / estorno

Decisione:
non hardcodare CFOP 1949 come recusa/estorno automatico.

Ogni caso CFOP 1949 deve essere verificato con:
- XML;
- natOp;
- NFref;
- prodotti;
- valori;
- segni;
- collegamento con documento originario.

Caso noto:
NF 359680 con CFOP 1949 e NF collegata probabilmente 358780 non devono essere contabilizzate se confermate come coppia di recusa/storno bonificação sem efeito contábil.

Ma:
CFOP 1949 può rappresentare anche altra saída/operação.
Non generalizzare.

## NF mista venda/bonificação

Decisione:
la logica deve essere item-level.

Non assumere che venda/compra normale sia solo CFOP 5405.
Verificare i CFOP reali presenti negli XML.

Venda/compra può includere, per esempio, CFOP 5102 e altri.
La mappatura deve essere basata sui CFOP reali, non su hardcoding isolato.

## Investimenti / imobilizado

Decisione:
beni come refrigeradores, cervejeiras, equipamentos, placas/motherboard e simili non devono essere mandati automaticamente a costo.

Trattamento prudente:
- mantenere in REVISAR;
- oppure usare conta tecnica di revisione investimento, se disponibile;
- non classificare direttamente come costo senza conferma contabile;
- non classificare automaticamente come immobilizzato definitivo senza revisione.

## PATAXO / materiais de obra

Decisione:
PATAXO è caso sensibile.

Non creare alias ampio `pataxo`.

Possibile trattamento:
- regola per fornecedor può essere usata con cautela;
- materiali prediali specifici possono andare a manutenzione/obra se confermati;
- evitare regole troppo ampie su singoli termini come cabo, fio, tigre, material.

## Alias e regras

Decisione:
preferire alias specifici e sicuri.

Usare alias quando:
- descrizione specifica/esatta ricorre;
- classificazione è chiara;
- rischio di falso positivo è basso.

Usare palavras-chave/regra solo quando:
- criterio è più ampio ma ancora sicuro;
- ci sono test;
- non cattura prodotti ambigui.

Non creare alias/regole generiche come:
- molho;
- shine;
- dry;
- tigre;
- pataxo;
- placa;
- m/b isolato;
- refr isolato;
- plastica;
- frutti isolati;
- açúcar / sem açúcar / zero açúcar;
- papel / saco / material / cabo / fio senza qualificatore.

## Fórmula resumo executivo caixa

Formula consolidata:

Cassa finale calcolata =
Abertura caixa
+ Total vendas bruto
+ Contas recebidas
+ Gorjetas
+ Antecipações
+ Suprimento caixa
- Cartão crédito fonte externa
- Cartão débito fonte externa
- PIX fonte externa/banco
- Desconto
- Conta a pagar
- Retirada

Interpretazione:
- `Total vendas` è valore bruto.
- `Gorjetas` sono campo operativo extra, non già incluse in Total vendas.
- `Antecipações` sono positive solo se entrano realmente nella cassa/fechamento e non sono già incluse in altre voci.
- `Retirada` include uscite fisiche di cassa, incluse sangrias e despesas pagas em dinheiro, salvo esposizione separata.
- Divergenza = cassa finale calcolata - cassa finale contata.
- Divergenza positiva = FALTA.
- Divergenza negativa = SOBRA.

## Estratti bancari e comprovanti

Decisione per Fase 13:
ogni mese ci saranno due estratti separati, uno per ciascun CNPJ.

Regola:
- import separato per CNPJ/conta;
- normalizzazione separata;
- consolidamento solo dopo normalizzazione;
- mantenere origine file e audit trail.

Gli estratti/comprovanti servono a rafforzare riconciliazione e DRE gerencial, non a trasformare automaticamente tutto in contabilità fiscale.

## Run e fonti

Decisione permanente:
quando cambiano regras, alias, parser, logica fiscale o contratto contabile, generare nuova run pulita da XML/fonte primaria.

Le run vecchie servono per:
- confronto;
- audit storico;
- regressione dati.

Non servono come base operativa se la logica è cambiata.
