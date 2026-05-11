# ⚡ Oscilogram Generator — Web App

Web verzija desktop alata za generiranje oscilogram grafova i Word izvještaja iz mjernih `.txt` fajlova.

## Pokretanje lokalno

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy na Streamlit Cloud (besplatno)

1. Pushaj ovaj folder na GitHub repo
2. Idi na [share.streamlit.io](https://share.streamlit.io)
3. Poveži GitHub repo → odaberi `app.py`
4. Klikni **Deploy** — gotovo!

## Struktura projekta

```
oscilogram_web/
├── app.py                  ← Streamlit sučelje
├── oscilogram_core.py      ← Originalna logika (nepromijenjena)
├── requirements.txt
└── README.md
```

## Workflow

1. **Tab ①** — Upload Excel/CSV + TXT fajlovi + Word template
2. **Tab ②** — Preview oscilogram s podešavanjem X/Y osi po komadu
3. **Tab ③** — Generiranje i download Word dokumenta (.docx) ili ZIP

## PRO funkcije

PRO lozinka se postavlja kao environment varijabla:

```
OSCGEN_PRO_PASSWORD=tvoja_lozinka
```

Na Streamlit Cloud: Settings → Secrets:
```toml
OSCGEN_PRO_PASSWORD = "tvoja_lozinka"
```

## Imenovanje TXT fajlova

Fajlovi moraju biti imenovani kao `v<No>.txt` ili `c<No>.txt`  
gdje `<No>` odgovara broju u Excel tablici.

Primjer: `v1.txt`, `c1.txt`, `v2.txt`, `c2.txt` …
