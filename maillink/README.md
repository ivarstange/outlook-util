# Mail Link – Outlook-tillegg

Viser og kopierer en direktelenke (deeplink) til den valgte eposten. Fungerer i nye Outlook, Outlook på nett og klassisk Outlook.

## Filer
- `manifest.xml` – tilleggets manifest (det er denne du sideloader)
- `taskpane.html` / `taskpane.js` – sidepanelet
- `assets/` – ikoner

## 1. Hosting (GitHub Pages)
Filene ligger i `maillink/` i repoet `ivarstange/outlook-util` og serveres fra
`https://ivarstange.github.io/outlook-util/maillink`.

Slå på Pages: Settings → Pages → Source: *Deploy from a branch* → branch `main`, mappe `/ (root)`.
(Pages på private repoer krever GitHub Pro/Team – ellers må repoet være offentlig.)

## 2. Manifestet
`manifest.xml` peker allerede på adressen over. Bytter du host, søk/erstatt
`https://ivarstange.github.io/outlook-util/maillink` og oppdater `<AppDomain>` (kun domenet).
Øk `<Version>` når du endrer manifestet, og sideload på nytt.

## 3. Sideload
- Gå til https://aka.ms/olksideload (eller Outlook på nett → Hent tillegg → Mine tillegg → Egendefinerte tillegg → Legg til fra fil) og velg `manifest.xml`.
- Jobbkonto: hvis det er sperret, må en M365-admin laste opp manifestet under Admin center → Innstillinger → Integrerte apper.

## Bruk
Åpne en epost → trykk **Hent lenke** på verktøylinjen (i nye Outlook finnes den også under «…»). Trykk på nåla i panelet for å holde det åpent mens du bytter epost.
Knapper: **Kopier lenke**, **Markdown** (`[emne](url)`), **Lenke med tittel** (rik tekst for Word/OneNote/Teams), **Åpne**.

## Lenkeformat
`https://outlook.office.com/owa/?ItemID=<REST-ID>&exvsurl=1&viewmodel=ReadMessageItem`
(`outlook.live.com` for private kontoer). Lenken åpner eposten i Outlook på nett for den som har tilgang til postboksen.

## Begrensninger
- Kun lesemodus (mottatte og sendte eposter). Utkast/nye meldinger har ingen ID før de er lagret.
- Delte postbokser kan kreve at du er logget inn med tilgang til den.
