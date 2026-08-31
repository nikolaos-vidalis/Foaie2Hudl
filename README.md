![Foaie2Hudl](assets/frf-to-hudl.png)

# Foaie2Hudl

Un instrument automatizat creat pentru antrenorii și analiștii de performanță din fotbalul juvenil din România, cu scopul de a simplifica munca administrativă de după meci.

## Descriere
Conversia foilor oficiale de arbitraj (din format `.pdf`) în documente structurate este un proces manual și de durată. **Foaie2Hudl** extrage automat datele din raportul arbitrului — competiția, data, rezultatul, numele echipelor, loturile de jucători și numerele de tricou — și completează șablonul de foaie de echipă Hudl/Wyscout, gata de descărcat ca `.docx`.

### Funcționalități principale
* **Extragere automată:** Procesează rapoartele `.pdf` ale arbitrilor pentru a colecta titularii, rezervele, numerele de pe tricou și datele despre meci.
* **Procesare în lot:** Poți încărca mai multe rapoarte simultan — fiecare este convertit separat, cu descărcare individuală sau a tuturor într-o arhivă `.zip`.
* **Interfață bilingvă:** Română (implicit) și engleză, comutabile din interfață. Doar site-ul este bilingv — documentul generat rămâne identic cu șablonul.
* **Formatat pentru Hudl/Wyscout:** Completează direct șablonul oficial, păstrând structura, formatarea și etichetele acestuia, cu textul în negru.
* **Adaptat pentru România:** Optimizat pentru rapoartele de meci utilizate în fotbalul românesc (formatele FRF / Football Connect), cu păstrarea diacriticelor.
* **Economie de timp:** Elimină introducerea manuală a datelor și erorile de redactare după fiecare etapă.

## Ce fișier se încarcă

Aplicația așteaptă **raportul complet de arbitraj** (`Raport Arbitru`, `.pdf`), descărcat din [www.footballconnect.ro](https://www.footballconnect.ro).

Acesta este documentul care conține foile de meci (`Foaie de meci`) ale **ambelor** echipe — gazde și oaspeți. **Foaia de meci a unei singure echipe nu este suficientă**, deoarece șablonul Hudl are nevoie de ambele loturi. Dacă încarci un astfel de fișier, aplicația te avertizează explicit.

Limita de încărcare este de **5 MB** per fișier (un raport obișnuit are câteva sute de KB).

## Utilizare

Aplicația este disponibilă online, gata de folosit — nu este nevoie de nicio instalare:

### 👉 [foaie2hudl.streamlit.app](https://foaie2hudl.streamlit.app/)

Interfața are trei panouri: în **stânga** încarci rapoartele, la **mijloc** este identitatea aplicației și comutatorul de limbă (română / engleză), în **dreapta** apar foile de echipă generate.

Încarcă unul sau mai multe rapoarte de arbitraj, verifică rezumatul afișat pentru fiecare meci și descarcă foile de echipă — individual sau toate într-o arhivă `.zip`.

Fișierele generate sunt denumite după competiție, dată și echipe:

```
Liga_Elitelor_U17_2026-08-30_SC_Dinamo_1948_vs_FC_Voluntari.docx
```

Documentul generat nu este niciodată tradus: conține etichetele proprii ale șablonului (în engleză) și datele extrase din raport, exact ca șablonul original.

> Șablonul `Wyscout teamsheet template.docx` trebuie să rămână în directorul aplicației — de acolo este citit la fiecare generare. Șablonul nu este modificat niciodată.

## Structura proiectului
| Fișier | Rol |
| --- | --- |
| `app.py` | Interfața Streamlit (încărcare, texte RO/EN, descărcare) |
| `parse_report.py` | Extrage informațiile meciului și loturile din PDF |
| `fill_teamsheet.py` | Completează șablonul `.docx` |

-----

> **Confidențialitatea datelor:**
> Aplicația Streamlit nu stochează și nu păstrează niciun fel de date sau fișiere încărcate. Procesarea fișierelor se realizează exclusiv temporar, în memorie, pe durata sesiunii curente. Aceeași explicație este afișată și în aplicație, sub zona de încărcare.
>
> **Data Privacy:**
> *The Streamlit application does not retain or store any uploaded data or files. All processing is carried out temporarily in memory for the duration of the active session. The same explanation is shown inside the app, below the upload area.*
