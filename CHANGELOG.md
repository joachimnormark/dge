# Changelog - Single Region App

## v11.2 - Tabel 11 tilføjet (Latest)

### ✨ Ny feature:

**Tabel 11: Grupper med udelukkende SGE-modul møder**
- Viser aktive grupper hvor ALLE godkendte møder er SGE-modul
- Expandable format - klik for at se detaljer
- Viser:
  - Gruppenavn
  - Antal SGE-modul møder
  - Mødetitler på de enkelte møder
- Kun i web-interface (ikke i PDF downloads)

### 📍 Placering:
- Efter Tabel 10 (Lukkede grupper)
- Før PDF Download sektionen

### 🔍 Logik:
En gruppe vises hvis:
1. Gruppen er aktiv (Status = 'Aktiv' eller ikke arkiveret)
2. Har mindst 1 godkendt møde i perioden
3. ALLE godkendte møder indeholder "SGE" og "modul" i Mødetype
4. Viser unique mødetitler

### 💡 Eksempel:
```
▶ 12-mandsgruppe Nord (3 møder)
   Mødetitler:
   • Introduktion til SGE-modul
   • SGE-modul efterårssemesteret
```

---

## v11.1 - Bug fix

### Fixed:
- ArrowDtype håndtering i filter_members_gruppeledere()

---

## v11.0 - Gruppestørrelse i PDF detaljer

### Features:
- Del 4 i PDF med detaljer: Grupper fordelt efter antal medlemmer

---

## v10.0 - Periode-labels, Individuelle søjler

### Features:
- Periode-labels i stedet for P1/P2
- Tabel 4: Individuelle søjler (1,2,3...14,15+)
- Tabel 5: Individuelle søjler (1,2,3...14,15+)
- Tabel 6: Opdateret beskrivelse (godkendte møder)

---

## Tidligere versioner

Se tidligere changelog for komplet historik.
