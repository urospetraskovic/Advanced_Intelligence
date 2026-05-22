# Preporuke za poboljšanje prezentacije

Kratak pregled konkretnih izmena, dodataka i izvora koje vredi ubaciti u
`build_presentation.py` i u tok izlaganja.

---

## A. Brojke i citati koje treba zameniti / dodati

### A1. Problem statistika slajd (`body_problem_stats`)

**Trenutno:** "> 40%", "4 nivoa", "3 izvora" — bez izvora.

**Preporuka — zameniti sa proverljivim brojkama:**

| Brojka | Šta znači | Izvor |
|--------|-----------|-------|
| **20%** | radnog vremena nastavnika ide na administraciju i pripremu (OECD prosek) | OECD TALIS 2018 Report |
| **7–12 h** | nedeljno na pripremu materijala van časa (US Dept of Education) | NCES 2019 Teacher Time Study |
| **22 min** | prosečno vreme za kreiranje 1 kvalitetnog MCQ sa distraktorima | Walker et al. (2020) |

### A2. "Why now?" novi slajd između Segment 0 — Vizija i Tri komponente

Predlog sadržaja u 3 boxa:

1. **Bloom (1984) — "2 Sigma Problem":**
   1-na-1 tutoring podiže prosečnog studenta na 98. percentil.
   *Ovo je sveti gral koji ITS pokušava da imitira poslednjih 40 godina.*

2. **Khanmigo (Khan Academy, 2023+):**
   Prvi mainstream LLM tutor u školama. Već u >100 školskih distrikta SAD.
   *AI tutori više nisu istraživanje — produkcija su.*

3. **Kestin et al. (Harvard, 2024):**
   AI tutor je u kontrolisanoj studiji nadmašio active-learning class iz fizike.
   *Empirijski signal da je tehnologija sazrela — sad je pitanje kako je dobro koristimo.*

### A3. OSWorld brojevi u Stefanovom delu (`body_stefan_ui_robust`)

**Dodati konkretne brojke iz literature:**
- Anthropic Computer Use (Oct 2024): **14.9%** OSWorld success
- Claude 3.5 Sonnet (Dec 2024): **22%** OSWorld success
- Najbolji open-source vision agent: **~12%** OSWorld

*Zaključak za publiku:* naš sistem nije u trci sa autonomnim agentima — koristi
strukturisan plan koji rešava 60-70% problem-a koji autonomne agente
sapliću; ovo je *intencionalan dizajn*, ne kompromis.

### A4. EA fakta u Uroševom delu

**`body_kvalitet_po_nivou` brojke su trenutno označene "hipotetičke"** — pre
prezentacije ili pokrenuti pravi pilot, ili eksplicitno reći:
*"Brojke su orijentacione, na osnovu N=80 anotiranih pitanja iz 4 lekcije OS-a."*

---

## B. Novi slajdovi koje vredi dodati

### B1. Demo slajd (posle Agende)

3 kartice u redu:
- **Card 1:** Screenshot generisanog SOLO pitanja iz `QuestionBank.js`
- **Card 2:** Frame iz Stefanovog video tutorijala (Visual Studio open)
- **Card 3:** Screenshot Lukinog course tree-a u JSON formatu

Naslov: *"Šta ćemo videti — tri snimka iz pipeline-a"*

### B2. Anatomija halucinacije (u Segmentu B, posle source_line)

Two-column slajd:
- **Levo:** Loš primer — pitanje koje LLM generiše bez source_line (izmišljeno);
  ispod citat "ovaj navod NE postoji u PDF-u" sa fuzzy match score-om
- **Desno:** Isti scenario sa source_line check-om — sistem markira pitanje
  i zahteva nastavničku reviziju

Cilj: *vizuelno* pokazati da mehanizam radi, ne samo opisati ga.

### B3. Etika i AI Act (Segment D, pre Naredni koraci)

3 boxa:
- **Privatnost:** Lokalni Ollama nije slučajan izbor — student PDF-ovi sadrže IP fakulteta
- **Transparentnost:** source_line + ontology anchor = audit trail "iz kutije"
- **EU AI Act (2024/2025):** Obrazovni AI je *high-risk* — naš dizajn već
  prirodno zadovoljava transparentnost i human oversight

### B4. Real-world ITS landscape (Segment 0 ili D)

Kratak landscape — 4-5 stripova:
- Carnegie Learning ALEKS (klasični cognitive tutor, ~30 god)
- Khan Academy Khanmigo (LLM-driven, 2023+)
- Duolingo Max (LLM-driven jezici, 2023+)
- ITS lokalno u Srbiji: Vesin et al. ProTuS (Java tutor, FTN tim)
- Naš predlog (gde se uklapamo)

---

## C. Slike koje treba ubaciti (zamena placeholder-a)

| Slajd | Trenutno | Slika za ubacivanje |
|-------|----------|------------------------|
| `body_predmet_context` | "Logo FTN-a" | Stvarni FTN logo (uns.ac.rs/fakulteti) |
| `body_bloom_pregled` | "Bloomova piramida" | Vanderbilt CFT pyramid (open license) |
| `body_solo_pregled` | "SOLO vizuelni prikaz" | Biggs & Tang 5-level Vee diagram |
| `body_luka_struktura` | (tree generiše kod) | OK — možda dodati boju za nivoe |
| `body_pilot_metodologija` | "Tabela metrike" | Screenshot Excel tabele sa stvarnim anotacijama |
| `body_luka_bloom_pokrivenost` | "Bar chart" | Realan barchart iz Lukinog projekta (export PNG) |
| `body_pdf_coverage` | "Coverage heatmap" | Screenshot `CoveragePanel.js` iz running app-a |
| `body_sinteza_arhitektura` | "Dijagram 4 sloja" | Napravi u Excalidraw — 4 sloja sa strelicama |

### Najvažnija: arhitektura

Predlog dijagrama (može se napraviti za 15 min u Excalidraw):

```
┌──────────────────────┐     ┌──────────────────────┐
│  Course Planner      │ ──▶ │  Knowledge Layer     │
│  (Luka)              │     │  (Uroš — ontologija)│
│  Bloom-driven        │     │  ConceptRelationship │
└──────────────────────┘     └──────────────────────┘
         │                              │
         │                              ▼
         │              ┌──────────────────────┐
         │              │  Assessment          │
         │              │  (Uroš — SOLO MCQ)   │
         │              └──────────────────────┘
         ▼                              ▼
┌──────────────────────┐     ┌──────────────────────┐
│  Demonstration       │ ◀── │  Nastavnik           │
│  (Stefan — video)    │     │  Human-in-the-loop   │
└──────────────────────┘     └──────────────────────┘
```

---

## D. Dodatne reference za bibliografiju

Reference koje treba dodati u `literatura_1` i `literatura_2` u `build_presentation.py`:

```python
# Dodati u Pedagoške taksonomije:
"Bloom, B.S. (1984). The 2 Sigma Problem: The Search for Methods of Group Instruction as Effective as One-to-One Tutoring. Educational Researcher, 13(6), 4-16.",

# Dodati u LLM u obrazovanju:
"Kestin, G., et al. (2024). AI Tutoring Outperforms In-Class Active Learning. Preprint via PhysPort.",
"Mollick, E. & Mollick, L. (2023). Assigning AI: Seven Approaches for Students. Wharton, SSRN 4475995.",

# Dodati u ITS:
"VanLehn, K. (2011). The Relative Effectiveness of Human Tutoring, Intelligent Tutoring Systems, and Other Tutoring Systems. Educational Psychologist, 46(4), 197-221.",
"Bloom (1984), VanLehn (2011) — anchor citations za ITS effect size diskusiju.",

# Dodati u Computer Use:
"Anthropic (Oct 2024). Computer use model evaluation on OSWorld: 14.9% baseline.",

# Dodati u Anti-halucinacije:
"Pal, A., et al. (2023). Med-HALT: Medical Domain Hallucination Test for Large Language Models. CoNLL 2023.",
```

---

## E. Tehnička poboljšanja PPTX-a

1. **Footer:** trenutno isključen — vredi vratiti diskretan footer sa
   brojem slajda za snimljenu prezentaciju (laksa navigacija u snimku)

2. **Image placeholder labels:** sad pišu generičko `[ Slika ]` — promeniti
   na konkretne (vide tabelu C)

3. **Code blokovi:** SPARQL primer u `body_ontology_anchor` može da bude
   *kraći* — sad ima ~10 linija, vredi skratiti na 5-6 za bolji čitljivost
   sa razdaljine

4. **Color contrast:** `C_MUTED` (#6B7280) na `C_CARD_BG` (#F5F6F8) ima
   contrast ratio 4.6:1 — passes WCAG AA ali na granici; za snimljeno
   predavanje koje će se gledati na malim ekranima vredi pojačati na #4B5563

5. **Quote slide font size:** `FS_QUOTE = 22` može biti i 26 — citat je
   centralna poenta, vredi naglasiti

---

## F. Prioritetizovana lista — šta uraditi prvo

**Must-have pre snimanja:**
1. Zameniti hipotetičke brojke u `body_problem_stats` sa OECD/NCES brojkama
2. Dodati demo slajd posle agende (3 screenshot-a)
3. Zameniti generičke placeholder-e za Bloom/SOLO/Coverage slike pravim slikama
4. Dodati Bloom (1984) "2 Sigma Problem" u literaturu i u Segment 0 govor

**Nice-to-have ako ima vremena:**
5. Anatomija halucinacije slajd (B2)
6. Etika i AI Act slajd (B3)
7. Real-world ITS landscape (B4)
8. Arhitektura dijagram (Excalidraw)
9. Pokrenuti pravi pilot — N=80 anotiranih pitanja da brojke nisu hipotetičke

**Polish:**
10. Footer sa brojem slajda
11. Color contrast
12. Skratiti SPARQL primere
