# Tabel 11: Grupper med udelukkende SGE-modul møder

## 📋 Hvad viser den?

**Tabel 11** identificerer aktive grupper hvor **ALLE** godkendte møder i den valgte periode er af typen "SGE-modul".

## 🎯 Formål

Find grupper der udelukkende fokuserer på SGE-modul aktiviteter - ingen andre mødetyper i perioden.

## 📊 Visning

For hver gruppe vises:
- **Gruppenavn**
- **Antal SGE-modul møder** i perioden
- **Mødetitler** - alle unikke titler på SGE-modul møderne

**Format:** Expandable (fold ud for at se detaljer)

## ✅ Kriterierne

En gruppe vises kun hvis:
1. ✅ Gruppen er **aktiv** (ikke arkiveret ELLER Status = 'Aktiv')
2. ✅ Gruppen har mindst **1 godkendt møde** i perioden
3. ✅ **ALLE** godkendte møder er af typen "SGE-modul"
4. ✅ Mødetype indeholder både "SGE" og "modul" (case-insensitive)

## 🔍 Hvad tæller som SGE-modul?

Møder hvor `Mødetype` indeholder:
- "SGE" (case-insensitive)
- OG "modul" (case-insensitive)

**Eksempler:**
- ✅ "SGE-modul"
- ✅ "SGE Modul"
- ✅ "sge-modul møde"
- ❌ "SGE møde" (mangler "modul")
- ❌ "Modul" (mangler "SGE")
- ❌ "DGE-møde" (forkert type)

## 💡 Eksempel output

```
▶ 12-mandsgruppe Nord (3 møder)
   Mødetitler:
   • Introduktion til SGE-modul
   • SGE-modul efterårssemesteret
   • Afslutning SGE-modul

▶ Supervision Gruppe A (2 møder)
   Mødetitler:
   • SGE-modul del 1
   • SGE-modul del 2
```

## ⚠️ Vigtigt

**ALLE godkendte møder skal være SGE-modul:**
- Hvis gruppen har holdt **blandet** møder (f.eks. både DGE-møde og SGE-modul) → vises IKKE
- Hvis gruppen har holdt **kun SGE-modul** møder → vises ✅

**Kun godkendte møder tæller:**
- Afviste møder ignoreres
- Afventende møder ignoreres
- Kun "Godkendt" status tæller

## 📍 Placering

- **I app:** Vist som Tabel 11 (efter Tabel 10)
- **I PDF:** ❌ IKKE inkluderet (kun i web-interface)

## 🔧 Tekniske detaljer

**Data krav:**
- `Gruppenavn` - til identifikation
- `Status` - til aktiv/inaktiv check
- `Mødetype` - til SGE-modul identificering
- `Mødetitel` - til visning af mødetitler
- Møder skal være filtreret til `Status = 'Godkendt'`

**Logik:**
```python
# Pseudo-kode
for hver gruppe:
    if gruppe er aktiv:
        mødetyper = alle godkendte møder for gruppen
        if alle mødetyper indeholder "sge" og "modul":
            vis gruppe med detaljer
```

## 🎯 Brug cases

1. **Identificer SGE-fokuserede grupper**
   - Find grupper der kun arbejder med SGE-modul

2. **Overvåg specialiserede aktiviteter**
   - Track dedikerede SGE-modul indsatser

3. **Kvalitetssikring**
   - Verificer at SGE-modul grupper ikke blander mødetyper

4. **Rapportering**
   - Dokumentér SGE-modul engagement

## ❓ FAQ

**Q: Hvorfor vises min gruppe ikke?**
A: Tjek at:
- Gruppen er aktiv
- ALLE godkendte møder er SGE-modul (ingen andre typer)
- Mødetype indeholder både "SGE" og "modul"

**Q: Kan en gruppe have andre møder der ikke er godkendt?**
A: Ja! Kun godkendte møder tæller. Afviste/afventende møder ignoreres.

**Q: Hvad hvis der er ingen grupper?**
A: Viser: "Ingen grupper har udelukkende holdt SGE-modul møder i perioden"

**Q: Hvorfor er dette ikke i PDF?**
A: Tabel 11 er designet til interaktiv visning (expandable format). PDF'er viser kun grafer og primære tabeller.
