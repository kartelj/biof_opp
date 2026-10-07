# Python za biologe: predavanja

Slajdovi na srpskom jeziku prema knjizi Martina Jonesa *Python for Biologists* (2013),
podeljeni na 9 lekcija, po jedna za svako poglavlje knjige. Svaka lekcija postoji kao PDF
(za čitanje i deljenje) i kao PPTX (za izmene i prikazivanje u PowerPoint-u).

## Lekcije

| # | Lekcija | Slajdova | PDF | PPTX |
|---|---------|---------:|-----|------|
| 1 | Uvod i radno okruženje | 9 | [PDF](lekcije/01-uvod-i-radno-okruzenje.pdf) | [PPTX](lekcije/01-uvod-i-radno-okruzenje.pptx) |
| 2 | Štampanje i rad sa tekstom | 26 | [PDF](lekcije/02-stampanje-i-rad-sa-tekstom.pdf) | [PPTX](lekcije/02-stampanje-i-rad-sa-tekstom.pptx) |
| 3 | Čitanje i pisanje fajlova | 18 | [PDF](lekcije/03-citanje-i-pisanje-fajlova.pdf) | [PPTX](lekcije/03-citanje-i-pisanje-fajlova.pptx) |
| 4 | Liste i petlje | 17 | [PDF](lekcije/04-liste-i-petlje.pdf) | [PPTX](lekcije/04-liste-i-petlje.pptx) |
| 5 | Pisanje sopstvenih funkcija | 15 | [PDF](lekcije/05-pisanje-sopstvenih-funkcija.pdf) | [PPTX](lekcije/05-pisanje-sopstvenih-funkcija.pptx) |
| 6 | Uslovni testovi | 15 | [PDF](lekcije/06-uslovni-testovi.pdf) | [PPTX](lekcije/06-uslovni-testovi.pptx) |
| 7 | Regularni izrazi | 17 | [PDF](lekcije/07-regularni-izrazi.pdf) | [PPTX](lekcije/07-regularni-izrazi.pptx) |
| 8 | Rečnici | 15 | [PDF](lekcije/08-recnici.pdf) | [PPTX](lekcije/08-recnici.pptx) |
| 9 | Fajlovi, programi i korisnički unos | 20 | [PDF](lekcije/09-fajlovi-programi-i-korisnicki-unos.pdf) | [PPTX](lekcije/09-fajlovi-programi-i-korisnicki-unos.pptx) |

Lekcija 1 počinje naslovnim slajdom i pregledom svih poglavlja, a lekcija 9 se završava
pregledom naučenog i završnim slajdom.

## Sadržaj foldera

- `lekcije/`: gotovi PDF i PPTX fajlovi.
- `izvor/`: izvor slajdova u HTML-u. `deck.json` određuje redosled slajdova i sekcije
  (poglavlja), a `slides/*.html` sadrži po jedan slajd (1920 × 1080 px). Beleške za
  predavača su u elementu `<aside>` na kraju svakog slajda.
- `alati/izvoz.py`: skripta koja od izvora pravi lekcije.

## Kako ponovo napraviti PDF i PPTX

Posle izmene slajdova u `izvor/`:

```bash
pip install python-pptx playwright
python -m playwright install chromium
python prezentacija/alati/izvoz.py          # sve lekcije
python prezentacija/alati/izvoz.py 3 5      # samo lekcije 3 i 5
```

Za PDF je potreban internet, jer se fontovi IBM Plex Sans i IBM Plex Mono preuzimaju sa
Google Fonts. Skripta na kraju ispisuje elemente koji možda izlaze iz okvira ili sa slajda.

## Razlike između PDF i PPTX

- **PDF** izgleda isto kao originalni slajdovi, sa fontovima IBM Plex Sans i IBM Plex Mono.
- **PPTX** koristi Arial i Consolas, koji dolaze uz Office, pa izgleda isto na svakom
  računaru. Raspored je izmeren za te fontove. Svaki tekst je zaseban okvir koji se može
  menjati, a beleške za predavača su u beleškama slajda.

## Napomene o sadržaju

- Primeri su prilagođeni Pythonu 3 i provereni. Knjiga je pisana dok se još koristio
  Python 2; razlike su objašnjene na slajdu „Python 3 i ova knjiga” i u beleškama.
- Nekoliko grešaka iz rešenja u knjizi je ispravljeno na slajdovima, sa objašnjenjem u
  beleškama (granice egzona u vežbi sa intronima, poređenje `int(expression)` u vežbama sa
  `data.csv`, uslov `<=` pri razvrstavanju sekvenci, izlazi u vežbama sa pristupnim
  brojevima i dvostrukom digestijom).
- Izlazi u vežbama koje čitaju fajlove iz materijala uz knjigu (adapteri, egzoni,
  digestija, k-meri) preuzeti su iz knjige.

## Licenca

Knjiga *Python for Biologists* objavljena je pod licencom Creative Commons
Attribution-NonCommercial-ShareAlike 3.0. Ovi slajdovi su izvedeno delo i dele se pod
istom licencom (CC BY-NC-SA 3.0), uz navođenje izvora.
