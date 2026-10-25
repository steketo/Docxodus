# Il ramo `autorel` di questa copia

Questa copia di [JSv4/Docxodus](https://github.com/JSv4/Docxodus) serve ad AutoRel, che scrive con `redline` le
bozze `.docx` con le proposte come revisioni native di Word, senza Word (voce SCR del suo piano di sviluppo; risposta
683 del titolare: correzioni in una copia nostra, proposte anche all'originale con richieste di fusione).

## Regole del ramo

- In `.github/workflows/` c'è il solo `autorel-redline.yml`, che parte solo a mano e pubblica i binari
  `autorel-<versione dell'originale>.<n>`; i 13 flussi dell'originale sono tolti, perché sui commit e sui tag
  partirebbero (`publish.yml` riscrive i binari a ogni tag `v*`). `scripts/autorel/controlla-flussi.sh` lo controlla
  prima di ogni spinta, come gancio `pre-push` del clone di lavoro, e nel flusso.
- Si spingono solo il ramo `autorel` e i tag `autorel-*`, più i rami `fix/*` delle richieste di fusione all'originale
  (§2.B del disegno, risposte 683 e 688); nessun tag `v*` si spinge mai. Una versione pubblicata non si riscrive.

## Aggiornare la copia a una versione nuova dell'originale

1. `git fetch originale --tags`, poi `git rebase --onto v<nuova> v14.1.0 autorel`.
2. `sh scripts/autorel/controlla-flussi.sh`: un flusso nuovo dell'originale si toglie nel primo commit del ramo.
3. Le prove di `Docxodus.Tests`, poi il flusso `autorel-redline` con la versione `<nuova>.1`, poi la misura
   d'arrivo dell'engine di AutoRel (`engine/py/misura_docxodus.py`). Una correzione accolta dall'originale esce dal
   ramo.

## Correzioni

| correzione | prova | richiesta di fusione |
|---|---|---|
