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
| Rinumero dei rinvii consecutivi: i token del risultato del k-esimo campo hanno il suffisso `fld:k` nel diff del capoverso accoppiato coi campi uguali (`IrDiffTokenizer.cs`, `IrEditScriptBuilder.cs`) | `BodyFieldSpanDiffTests.Compare_ConsecutiveRefFieldsRenumbered_EachFieldKeepsItsResult` (rosso sul caso di 7 rinvii prima della correzione) | da aprire |
| Segnalibro scritto due volte: col capoverso cancellato per intero, la copia del «prima» (in `w:del`) restava accanto a quella del «dopo» nel capoverso inserito o in uno uguale, e Word, accettato tutto, stampava «Errore. L'origine riferimento non è stata trovata.» nei `REF`; ora resta la sola copia del «dopo», tolta dal `w:ins` quando sta in un capoverso (`IrMarkupRenderer.CollapseWholeBlockBookmark`, chiamata da `NormalizeBookmarks`) | `DocxDiffBookmarkStructureTests.Bookmark_Compare_HasUniqueBookmarkNames` (rosso prima della correzione su `whole-para-rewrite` e sul caso nuovo `bookmark-moves-to-equal-para-from-deleted-para`) | da aprire |
| Segnalibro perso rifiutando (14.1.0.3): una figura tolta e la didascalia seguente rinumerata, il segnalibro comune su «Figura 24» restava solo sul testo inserito, e Word, rifiutato tutto, lo toglieva col testo; ora, come nel confronto di Word, «Figura 24» della didascalia cancellata resta testo comune col segnalibro, il segno di capoverso resta cancellato, ed etichetta e numero della didascalia seguente sono cancellati (`IrMarkupRenderer.CarryBookmarksOntoCommonText`, passo (A0) di `NormalizeBookmarks`) | `DocxDiffRenumberedCaptionBookmarkTests.Compare_BookmarkOnARenumberedCaption_EnclosesTheSameTextInBothViews` (rossa prima della correzione sui 3 casi: rifiutando il segnalibro non racchiudeva più testo), coi controlli `…_LabelFormattingChanged_…`, `…_ParagraphPropertiesDiffer_…` e `Compare_SecondBookmarkFromTheSameStart_KeepsTheEarlierHandling`; dopo la revisione avversaria anche `Compare_BookmarksOnCaptionsRenumberedInCascade_…` e le prove `Pass_…` sul markup scritto a mano (candidati incrociati, spostamento, marcatori estranei); al terzo giro i segnalibri nascosti non citati (`Compare_HiddenBookmarkOutsideTheLabel_StillMovesTheBookmark`) e gli estremi accoppiati per ordine; al quarto giro l'eccezione ristretta a `_GoBack` e `_Hlk…` con l'intervallo da una parte della copia, e i nomi citati letti in tutte le storie (`Compare_RefBookmarkOnTheWholeCaptionCitedOnlyByAFootnote_…`, `Compare_HiddenBookmarkCitedOnlyByAFootnote_…`) | da aprire |
