#!/bin/sh
# Sul ramo autorel di steketo/Docxodus .github/workflows/ porta il solo autorel-redline.yml: un flusso dell'originale
# partirebbe a ogni spinta o tag e, coi tag v*, riscriverebbe i binari delle versioni (publish.yml:3-6, :467).
# Senza argomenti controlla HEAD (il primo lavoro del flusso, il commit dei flussi). Come gancio pre-push del clone
# di lavoro git gli passa il remoto e l'indirizzo, e su stdin una riga per riferimento spinto: rifiuta ogni tag v*, e
# del solo ramo autorel controlla l'albero del commit spinto, non l'albero di lavoro (revisione del piano, giro 1,
# rilievo 12). Un ramo di richiesta di fusione (fix/*) nasce da originale/main e porta i 13 flussi dell'originale: lì
# il controllo rifiuterebbe ogni spinta (giro 3, rilievo 1), e nessuno di quei flussi parte su una spinta a un ramo
# diverso da main. Ogni altro ramo, main compreso, si rifiuta: su main i 13 flussi partono a ogni spinta (giro 4,
# rilievo 1); le cancellazioni passano.
set -e
cd "$(git rev-parse --show-toplevel)"
controlla() {
  altri=$(git ls-tree --name-only "$1" .github/workflows/ | sed 's#.*/##' | grep -vx 'autorel-redline.yml' || true)
  if [ -n "$altri" ]; then
    echo "flussi dell'originale in $1, da togliere prima di spingere:" $altri >&2
    exit 1
  fi
}
if [ $# -eq 0 ]; then
  controlla HEAD
  exit 0
fi
while read -r locale sha remoto _; do
  case "$remoto" in
    refs/tags/v*) echo "un tag v* non si spinge sulla copia: $remoto (publish.yml dell'originale)" >&2; exit 1 ;;
  esac
  # Una cancellazione (sha di soli zeri) passa: anche --delete main di spingi.sh (Task 4, passo 4).
  case "$sha" in
    *[!0]*) ;;
    *) continue ;;
  esac
  # Elenco chiuso dei rami (giro 4, rilievo 1): main porta i 13 flussi, che lì partono a ogni spinta.
  case "$remoto" in
    refs/heads/autorel) controlla "$sha" ;;
    refs/heads/fix/*) ;;
    refs/heads/*) echo "sulla copia si spingono solo autorel e i rami fix/*: $remoto" >&2; exit 1 ;;
  esac
done
