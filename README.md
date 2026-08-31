![Foaie2Hudl](assets/frf-to-hudl.png)

# Foaie2Hudl

Un instrument automatizat creat pentru antrenorii și analiștii de performanță din fotbalul juvenil din România, cu scopul de a simplifica munca administrativă de după meci.

## Descriere
Conversia foilor oficiale de arbitraj (din format `.pdf`) în documente structurate este un proces manual și de durată. **Foaie2Hudl** extrage automat datele din raportul arbitrului — competiția, data, rezultatul, numele echipelor, loturile de jucători și numerele de tricou — și completează șablonul de foaie de echipă Hudl/Wyscout, gata de descărcat ca `.docx`.

### Funcționalități principale
* **Extragere automată:** Procesează rapoartele `.pdf` ale arbitrilor pentru a colecta titularii, rezervele, numerele de pe tricou și datele despre meci.
* **Formatat pentru Hudl/Wyscout:** Completează direct șablonul oficial, păstrând structura și formatarea acestuia.
* **Adaptat pentru România:** Optimizat pentru rapoartele de meci utilizate în fotbalul românesc (formatele FRF / Football Connect), cu păstrarea diacriticelor.
* **Economie de timp:** Elimină introducerea manuală a datelor și erorile de redactare după fiecare etapă.

## Instalare
```bash
pip install -r requirements.txt
```

## Rulare
```bash
streamlit run app.py
```

Încarcă raportul de arbitraj în format `.pdf`, verifică rezumatul meciului afișat și descarcă foaia de echipă completată.

> Șablonul `Wyscout teamsheet template.docx` trebuie să rămână în directorul aplicației — de acolo este citit la fiecare generare.

## Structura proiectului
| Fișier | Rol |
| --- | --- |
| `app.py` | Interfața Streamlit (încărcare + descărcare) |
| `parse_report.py` | Extrage informațiile meciului și loturile din PDF |
| `fill_teamsheet.py` | Completează șablonul `.docx` |

-----

> **Confidențialitatea datelor:**
> Aplicația Streamlit nu stochează și nu păstrează niciun fel de date sau fișiere încărcate. Procesarea fișierelor se realizează exclusiv temporar, în memorie, pe durata sesiunii curente.
>
> **Data Privacy:**
> *The Streamlit application does not retain or store any uploaded data or files. All processing is carried out temporarily in memory for the duration of the active session.*
