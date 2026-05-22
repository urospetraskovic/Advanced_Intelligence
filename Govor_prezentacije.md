# Govor prezentacije — LLM-driven ITS pipeline

> Speaker script za snimljenu prezentaciju (60–90 min) na predmetu **Napredne tehnike računarske inteligencije**, master FTN UNS.
> Tim: Uroš Petrašković, Luka Šarić, Stefan Lazarević.
> Tempo: ~130 reči/min sa pauzama → ~10 000 reči = ~75 min.

Konvencije:
- **(N)** = slajd broj u redosledu PPTX-a (counter koji koristi `build_presentation.py`).
- **`Ko priča`** = predlog ko od članova tima vodi taj deo (rotira se).
- `[pauza]`, `[pokaži sliku X]`, `[demo]` = scenske napomene.

---

## (1) Naslovna — Uroš (~30 s)

**Uroš:**
Dobar dan. Mi smo Uroš Petrašković, Luka Šarić i Stefan Lazarević, studenti master studija na Fakultetu tehničkih nauka, smer Računarstvo. Na predmetu Napredne tehnike računarske inteligencije izlažemo temu *„Od ishoda učenja do automatizovanog Intelligent Tutoring System-a"* — kako kombinujemo velike jezičke modele sa formalnim ontologijama i pedagoškim taksonomijama da bismo napravili sistem koji od ishoda učenja, automatski, dolazi do snimljenog video tutorijala i kviza zasnovanog na SOLO taksonomiji.

`[pauza, prelaz na sledeći slajd]`

---

## (2) Tim — Uroš (~1 min)

**Uroš:**
Pre nego što uđemo u temu, kratko o nama. Tema koju izlažemo nastala je spajanjem tri pojedinačna projekta koja smo do sada radili u okviru predmeta SOTIS — Savremene Obrazovne Tehnologije i Standardi. Svaki od nas je u SOTIS-u radio jedan rad. Luka je radio na automatskoj generaciji strukture kursa zasnovane na Bloomovoj taksonomiji. Stefan je razvijao sistem koji od prirodno-jezičke instrukcije pravi snimljeni video tutorijal — sa OWL ontološkom validacijom plana. Ja sam radio SOLO Quiz Generator — sistem koji od PDF lekcije pravi pitanja na četiri SOLO nivoa, sa ontology grounding-om i citatima iz izvornog teksta.

Sva tri rada su komplementarna — adresiraju različite faze istog problema. Vredi pomenuti i da je jedan od ovih radova već prihvaćen na konferenciji Sinteza 2026 — što nas je dodatno motivisalo da ih objedinimo i pokažemo kao integrisanu celinu, a ne tri odvojena projekta.

`[prelaz]`

---

## (3) Predmet i kontekst — Uroš (~1 min)

**Uroš:**
Predmet Napredne tehnike računarske inteligencije pokriva temu na preseku dve oblasti sa zvanične liste tema: *AI in Computing Education* i *Intelligent Tutoring Systems*. Format polaganja je snimljena prezentacija u trajanju od jednog do jednog i po sata, što nam daje prostor da uđemo dovoljno duboko u tehnička rešenja, ali i da tematski kontekstualizujemo zašto je problem koji rešavamo vredan rešavanja.

Predmet drže profesorka Simona Prokić i profesor Aleksandar Kovačević, a naša prezentacija će pokušati da im — i vama, ako ovo gleda neko drugi — pruži ne samo informaciju o tri konkretna projekta, nego i pedagoški okvir oko kojeg smo ih objedinili.

`[prelaz]`

---

## (4) Agenda — Uroš (~2 min)

**Uroš:**
Prezentacija je organizovana **tematski, ne po projektima**. To je svesna odluka. Da smo svaki radio jedno predavanje o svom projektu, dobili biste tri male prezentacije lepljene jednu za drugu. Ovako, svaki segment otvara jednu temu — pedagoške taksonomije, ontologije, evaluaciju, sintezu — i kroz nju prolazi iz uglova sva tri projekta. Cilj je da prati ideju, ne autora.

Pet segmenata:

**Segment 0 — Uvod i motivacija, deset minuta.** Krećemo od konkretnog problema: koliko vremena nastavnik troši na pripremu materijala i šta sve tu može da ide loše. Pokazujemo viziju ITS pipeline-a i kako se tri komponente uklapaju.

**Segment A — LLM prompting i pedagoške taksonomije, dvadeset pet minuta.** Bloomova taksonomija, SOLO taksonomija, PS4 prompt template, hijerarhijska dekompozicija. Tradeoff prompt engineering vs fine-tuning, lokalni vs cloud modeli.

**Segment B — Ontologije i formalni grounding, dvadeset pet minuta.** OWL, RDF, SPARQL — taman koliko treba da pratite priču. Pa onda konkretne realizacije: kako svaki član tima koristi ontologiju drugačije, i kako se anti-halucinacioni mehanizmi razlikuju.

**Segment C — Evaluacija, kvalitet i ograničenja, dvadeset pet minuta.** Pilot evaluacija SOLO pitanja, Bloomova pokrivenost, robustnost izvršenja video tutorijala, PDF coverage tracking. Završavamo razgovorom o granicama automatizacije i mestu čoveka u petlji.

**Segment D — Sinteza, budući rad, zaključak, deset do petnaest minuta.** Predlog integrisane četvoroslojne arhitekture, veze sa Deep Reinforcement Learning-om i Recommender Systems-ima, otvorena pitanja, zaključak.

Ako bi neko zapamtio samo jednu stvar sa ove prezentacije, voleo bih da to bude ovo: **LLM sam nije dovoljan. LLM plus ontologija plus pedagoška taksonomija je više od zbira svojih delova.**

`[prelaz na Segment 0 divider]`

---

# SEGMENT 0 — UVOD I MOTIVACIJA

## (5) Divider Segment 0 — Stefan (~30 s)

**Stefan:**
Segment nula. Uvod i motivacija. Pre nego što pričamo o LLM-ovima, ontologijama, taksonomijama — moramo da odgovorimo na pitanje *zašto*. Zašto je ovaj problem vredan rešavanja, i zašto ga rešavamo sad.

`[prelaz]`

---

## (6) Problem — Stefan (~3 min)

**Stefan:**
Kreiranje obrazovnih materijala je vremenski i kognitivno zahtevan posao. Postoji konkretan podatak iz OECD TALIS istraživanja iz 2018: prosečan nastavnik u OECD zemljama provede oko dvadeset procenata radnog vremena na pripremu materijala i administraciju — to je svaka peta radna ura. Studije iz Sjedinjenih Država idu još dalje — između sedam i dvanaest sati nedeljno samo na pripremu, izvan časa.

`[pokaži tri brojke na ekranu]`

Konkretnije — Walker i kolege su 2020. godine merili koliko traje da jedan kvalitetan multiple-choice item dobije, sa svim distraktorima i pravilno mapiran na ishod učenja: prosek **dvadeset dva minuta** po pitanju. Pomnožite to sa kvizom od dvadeset pitanja po lekciji, sa petnaest lekcija po semestru — dolazimo do desetina sati samo na pitanja.

`[pauza]`

Šta nastavnik konkretno radi rukama? Prvo, planira strukturu kursa — module, koncepte, ishode učenja. Drugo, piše objašnjenja, primere, definicije — i pri tome ih usklađuje sa pretpostavljenim nivoom studenata. Treće, priprema praktične demonstracije, najčešće u formi video tutorijala. I četvrto, sastavlja pitanja na različitim kognitivnim nivoima — i tu nastaje uska grla. Jedno pitanje na nivou "ponovi definiciju" se piše za pet minuta. Jedno pitanje na nivou "primeni koncept na novi problem" se piše za pola sata ili sat.

Sve to mora da bude pedagoški dosledno — što znači, mora biti svesno toga kojem nivou kognitivne složenosti pripada svako pitanje, i mora pokrivati ishode učenja na uravnotežen način. To dodatno otežava proces.

To je problem. Sada — zašto sad?

`[prelaz]`

---

## (7) LLM otvara vrata — Stefan (~3 min)

**Stefan:**
Pojava velikih jezičkih modela — ChatGPT 2022, GPT-4 2023, lokalni modeli kao Llama i Qwen 2024 — otvorila je realnu mogućnost da se deo ovog posla automatizuje. **LLM može da generiše tekst iz prirodno-jezičkog uputstva, da razlaže kompleksne zadatke na korake, da iterira brzo, i da radi multimodalno — tekst, kod, čak i sliku.** Sve te sposobnosti su direktno relevantne za pripremu nastave.

I to nije samo teorija. Khan Academy je 2023. lansirao Khanmigo — LLM tutora koji sad već radi u stotinama škola u SAD. Duolingo Max je živ od 2023. Postoji studija koju vodi Greg Kestin sa Harvarda, objavljena 2024 — pokazala je da AI tutor zasnovan na GPT-4 može da nadmaši aktivno-učeničku nastavu iz fizike, i to za pola vremena.

`[pauza]`

Ali ovde je važan zaokret. **Naivna primena LLM-a uvodi novi skup problema.**

Prvi problem — **halucinacije**. Model generiše tekst koji zvuči dobro, ali sadrži činjenice koje nisu tačne, ili koje ne postoje u izvornom materijalu. Pal i kolege su 2023. izmerili da, bez RAG arhitekture, GPT-4 ima oko petnaest do dvadeset procenata stopu halucinacije na faktičkim pitanjima. To je previše za obrazovanje.

Drugi problem — **generičnost**. Pitate ga: "napravi pitanje iz operativnih sistema." Daje vam pitanje o operativnim sistemima — ali generičko, kao iz Wikipedije, bez veze sa konkretnim PDF-om koji ste vi koristili u nastavi. Studenti su čitali baš taj PDF, ne Wikipediju.

Treći problem — **pogrešan kognitivni nivo**. Tražite pitanje "po Bloomovom Analyze nivou". Daje vam pitanje koje je u suštini Remember-nivo, samo zamotano u dužu rečenicu. Kognitivni nivo se ne kontroliše imenom — kontroliše se strukturom.

Četvrti problem — **odsustvo formalne strukture**. Svaki run vam daje drugačiji format. JSON koji jedanput radi, drugi put se polomi. Da bi ovo bio sistem, ne anegdota, treba ponovljivost.

I peti — **transparentnost**. Kad vam model proizvede odgovor, vrlo je teško da rekonstruišete *gde se on oslonio na izvor*. Bez toga, ne možete da znate da li je naučio iz vaše lekcije ili je improvizovao.

Ovih pet problema — to je razlog zašto LLM sam nije dovoljan. I to nas vodi u našu viziju.

`[prelaz]`

---

## (8) Vizija pipeline-a — Luka (~3 min)

**Luka:**
Mi predlažemo end-to-end pipeline. Pet koraka, bez ručne intervencije između njih.

`[pokaži dijagram sa pet boxova]`

**Korak jedan — ulaz.** Tri vrste informacija. Lista ishoda učenja u prirodnom jeziku — ono što nastavnik napiše u silabusu. PDF materijal — slajdovi, knjige, lekcije, šta god ima. I prirodno-jezičke instrukcije za praktične korake — na primer „kreiraj C# konzolni projekat u Visual Studio-u koji ispisuje Hello World".

**Korak dva — struktura kursa.** Bloom-driven. Sistem razlaže kurs u stablo modula, koncepata, ishoda i aktivnosti. Svaka aktivnost se mapira na neki Bloomov nivo. Ovo je moj deo projekta.

**Korak tri — ontologija.** OWL, RDF — formalna reprezentacija konceptualnih relacija. Šta je preduslov čemu, šta je primer čega, šta zavisi od čega. Ovo je osnova za sve što sledi.

**Korak četiri — generisanje sadržaja.** SOLO pitanja iz PDF lekcije — to je Urošev deo. Video demonstracije iz prirodno-jezičke instrukcije — to je Stefanov deo. Oba se oslanjaju na ontologiju iz prethodnog koraka.

**Korak pet — izlaz.** Strukturisan kurs sa hijerarhijom modula i aktivnosti. Skup pitanja po SOLO nivoima sa multiple choice opcijama, ontološkim anchor-ima, i citatima iz materijala. Skup video tutorijala čiji je izvršni plan validiran OWL ontologijom.

`[pauza]`

Ono što je važno reći — ovo nije sistem koji zamenjuje nastavnika. Ovo je sistem koji nastavniku snima — koliko procena — između sedamdeset i osamdeset procenata mehaničkog rada. Definisanje ishoda, donošenje stručnih odluka, recenzija — ostaje nastavniku. Sistem snima sve između.

`[prelaz]`

---

## (9) Tri komponente — Luka (~3 min)

**Luka:**
Pošto je svaki član tima radio jedan deo, evo kratkog pregleda — ko šta radi i šta tačno vidite na ekranu kad otvorite svaki od ovih projekata.

`[pokaži tri kartice]`

**Moj projekat — Course Structure Generator.** Bloom-driven hijerarhija. Ulaz: tema kursa, lista ishoda učenja. LLM iterira po Bloomovim nivoima — Pamtiti, Razumeti, Primeniti, Analizirati, Evaluirati, Kreirati — i za svaki nivo generiše skup aktivnosti. Pamti šta je već generisao, da bi izbegao redundanciju. Izlaz: JSON koji se može serijalizovati u OWL.

**Stefanov projekat — Computer Use Video Instructions.** Prirodno-jezička instrukcija ulazi, MP4 video izlazi. Između — LLM razlaže zadatak u plan, plan se mapira u OWL ontologiju, ontologija se validira SPARQL upitima, validan plan se izvršava — i pri tome se ekran snima FFmpeg-om. Vizijski model (Qwen 2.5 VL preko OpenRouter-a) lokalizuje UI elemente za klikove.

**Urošev projekat — SOLO Quiz Generator.** PDF lekcija ulazi, multiple choice pitanja izlaze. Lokalni Ollama model (Qwen 2.5 14B). Pitanja su organizovana po četiri SOLO nivoa — Unistructural, Multistructural, Relational, Extended Abstract. Svako pitanje nosi *source_line* — doslovni navod iz PDF-a koji opravdava tačan odgovor. Relaciona i Extended Abstract pitanja su dodatno ankorisana na konkretnu vezu iz ontologije.

Sve tri komponente koriste ontologiju kao zajednički jezik. Sve tri primenjuju neki oblik pedagoške taksonomije. Sve tri imaju ugrađeni anti-halucinacioni mehanizam — drugačiji u svakoj.

`[prelaz]`

---

## (10) Ključna poruka uvoda — Uroš (~2 min)

**Uroš:**
Pre nego što uđemo u tehničke detalje, hoću da postavim okvir.

`[pokaži formula box]`

**LLM + ontologija + pedagoška taksonomija = pouzdano, pedagoški kalibrisano ITS okruženje.**

I još — *više od zbira svojih delova*.

LLM je generator. Sam — pravi tekst, ali bez strukture, bez kontrole, bez transparentnosti. Ontologija je struktura. Sama — daje vam grof tripleta, ali nikakvog sadržaja koji ide u taj grof. Pedagoška taksonomija je kompas. Sama — kaže vam *kako bi nešto trebalo da izgleda*, ali ništa ne pravi.

Zajedno, sa pravim spojnim tkivom — *imate sistem*. Sistem koji LLM koristi kao generator, ontologiju kao ugovor o tome šta sme da izađe, a taksonomiju kao kriterijum kvaliteta.

To je glavna teza. Ostatak prezentacije je tehnička razrada svake od ove tri komponente i njihove integracije.

`[prelaz u Segment A]`

---

# SEGMENT A — LLM PROMPTING I PEDAGOŠKE TAKSONOMIJE

## (11) Divider Segment A — Uroš (~30 s)

**Uroš:**
Segment A. LLM prompting i pedagoške taksonomije. Sledećih dvadeset pet minuta. Glavna teza: hijerarhijska dekompozicija plus formalna pedagoška taksonomija ugrađena u prompt transformišu LLM iz generičkog generatora teksta u pedagoški kalibrisan alat.

`[prelaz]`

---

## (12) Bloom — Uroš (~3 min)

**Uroš:**
Krećemo od Bloomove taksonomije. Originalna verzija je iz 1956. godine — Benjamin Bloom i kolege. Revidirana verzija — koju mi koristimo — je iz 2001. godine, Anderson i Krathwohl. Šest nivoa misaonih operacija, hijerarhijski organizovani.

`[pokaži 6 nivoa]`

**Pamtiti** — najniži nivo. Setiti se činjenice, definicije, terminologije. „Šta je TCP?"

**Razumeti** — objasniti vlastitim rečima, parafrazirati. „Objasni razliku između TCP i UDP-a."

**Primeniti** — rešiti analogni problem. „Date su parametri TCP konekcije — izračunaj propustni opseg."

**Analizirati** — razložiti, uporediti, povezati. „Zašto bi TCP bio loš izbor za real-time video streaming?"

**Evaluirati** — proceniti, kritikovati, doneti sud. „Argumentuj za i protiv upotrebe QUIC-a umesto TCP-a u savremenim web aplikacijama."

**Kreirati** — sintetisati, napraviti nešto novo. „Predloži novi mehanizam za congestion control."

`[pauza]`

Bloomova taksonomija je *najuticajniji okvir u obrazovanju u dvadesetom veku*. To nije preterivanje. Praktično svaki nacionalni kurikulum, svaki silabus, svaka skala procene se na neki način oslanja na nju.

Za nas, Bloom upravlja **šta** sistem treba da generiše — koje **tipove aktivnosti**. To je centralna stvar u Lukinom projektu. On ne pravi pitanja — on pravi *aktivnosti različitih tipova*. Pitanja su samo jedan tip aktivnosti. Postoje i čitanja, vežbe, projekti, diskusije, simulacije.

`[prelaz]`

---

## (13) SOLO — Uroš (~3 min)

**Uroš:**
SOLO taksonomija. *Structure of Observed Learning Outcomes*. Biggs i Collis, 1982. Razlika od Bloom-a je suptilna ali važna.

Bloom govori o **tipu misaone operacije**. SOLO govori o **strukturi odgovora studenta**. Koliko elemenata postoji u odgovoru, i kako su organizovani.

`[pokaži 4 nivoa]`

**Unistructural** — odgovor sadrži jedan izolovan element. „Šta čuva PCB?" — „PID." Jedna činjenica.

**Multistructural** — odgovor sadrži više elemenata, ali nije pokazana veza između njih. „Koja tri stanja procesa najčešće prepoznajemo?" — „Ready, Running, Blocked." Tri činjenice, ali nije objašnjeno kako se prelazi između njih.

**Relational** — odgovor pokazuje *vezu* između elemenata. „Kako kontekstno prebacivanje utiče na throughput sistema?" — student mora da poveže koncept context switch-a, koncept overhead-a, i koncept throughput-a — i da pokaže relaciju.

**Extended Abstract** — odgovor *generalizuje* — primenjuje koncept na novi kontekst, ili izvodi novi princip. „Kako biste prilagodili scheduling algoritam za real-time sistem sa zagarantovanim deadline-om?" — student mora da uzme znanje o scheduling-u u opštem smislu i da ga transferuje u specifični kontekst.

`[pauza]`

Postoji peti nivo koji se ne koristi često — *Prestructural* — odgovor koji ne pokazuje razumevanje uopšte. „Ne znam." Mi ga ne generišemo svesno.

Postoji jedan jako zanimljiv rad od Listera i kolega iz 2006 — *Not Seeing the Forest for the Trees: Novice Programmers and the SOLO Taxonomy.* Oni su pokazali da početnici u programiranju imaju ozbiljnih problema sa Relational i Extended Abstract zadacima — i to objašnjava puno toga što vidimo u nastavi programiranja. Nemaju mentalni model dovoljno strukturisan da povežu dve stvari u kodu.

Za nas, SOLO upravlja **kako duboko** sistem treba da proverava znanje. To je centralna stvar u mom delu projekta. Svako pitanje koje moj sistem generiše ima eksplicitan SOLO nivo, i prompt template se prilagođava nivou — što ćemo videti za par slajdova.

`[prelaz]`

---

## (14) Bloom vs SOLO — Uroš (~3 min)

**Uroš:**
Pitanje koje se prirodno postavlja — *zašto obe*?

`[pokaži tabelu]`

Bloom i SOLO nisu konkurentske taksonomije. Komplementarne su. Pogledajte u tabeli redom.

**Tip hijerarhije.** Bloom: hijerarhija tipova misaonih operacija. SOLO: hijerarhija strukturalne složenosti odgovora.

**Broj nivoa.** Bloom: 6. SOLO: 5 sa Prestructural, 4 sa kojima radimo.

**Osnovna jedinica.** Bloom: glagol akcije — *znati*, *primeniti*, *analizirati*. SOLO: broj i organizacija elemenata u odgovoru.

**Tipična upotreba.** Bloom se koristi pri *planiranju aktivnosti* — koje aktivnosti treba da budu u kursu da pokriju različite nivoe. SOLO se koristi pri *evaluaciji dubine učenja* — kad gledamo studentov odgovor, na kom je nivou.

**Mi koristimo Bloom za strukturu kursa, SOLO za generisanje pitanja.** To nije slučajno.

`[pauza]`

Da to potpuno razjasnim: **Bloom drives the WHAT — koje tipove aktivnosti pravimo. SOLO drives the HOW DEEP — koliko duboko ih proveravamo.**

Možete imati Remember-nivo aktivnost koja je SOLO Unistructural — „šta je PCB?". Možete imati Remember aktivnost koja je SOLO Multistructural — „nabroji elemente PCB-a." Različite dimenzije.

Lukin sistem proizvodi *aktivnosti različitih Bloomovih nivoa*. Moj sistem onda — za svaku aktivnost koja je tipa "kviz" — proizvodi pitanja na različitim SOLO nivoima. Vertikalna integracija dve taksonomije.

`[prelaz]`

---

## (15) Luka — struktura kursa — Luka (~4 min)

**Luka:**
Sad ja preuzimam, da pokažem kako moj sistem radi.

`[pokaži stablo levo, opis desno]`

Hijerarhijska dekompozicija. Pet nivoa: **kurs → modul → koncept → ishod učenja → aktivnost**.

Ulaz je jednostavan: tema kursa, opis, lista ciljeva — ishoda učenja — koje nastavnik želi da postigne.

Onda LLM iterira po nivoima. Prvi prolaz — generiši module iz teme. „Operativni sistemi" → moduli kao „Procesi i niti", „Upravljanje memorijom", „Fajl sistemi", „Sinhronizacija"...

Drugi prolaz — za svaki modul, generiši koncepte. „Procesi i niti" → koncepti kao „Proces", „Nit", „Kontekstno prebacivanje", „Sinhronizacioni primitivi"...

Treći prolaz — za svaki koncept, generiši ishode učenja. „Proces" → „Razumeti stanja procesa", „Naučiti strukturu PCB-a", „Razumeti context switching"...

Četvrti prolaz — za svaki ishod, generiši aktivnosti. „Razumeti stanja procesa" → aktivnosti kao „pročitati sekciju 3.2", „uraditi vežbu sa state-diagram-om", „položiti mini-kviz na nivou Razumeti"...

Svaka aktivnost dobija eksplicitan **Bloomov nivo**. To se radi kroz prompt — LLM se eksplicitno traži da generiše aktivnost koja proverava ciljani Bloom nivo.

`[pauza]`

Postoji jedan ključan trik. **Anti-redundancija.** Kad LLM generiše osmu aktivnost za isti modul, lako se ponavlja — daje vam pet varijanti istog. Da bismo to sprečili, ja u prompt ubacujem listu *prethodno generisanih aktivnosti* sa eksplicitnim uputstvom — „nemoj da ovo ponavljaš." To se zove **state-of-generation context**. LLM dobija ne samo trenutni zadatak, nego i sliku gde smo do sada stigli.

Rezultat: koherentnost kroz ceo kurs. Nijedna aktivnost se ne ponavlja, i pokrivenost Bloomovih nivoa je ravnomerna — što ćemo izmeriti u Segmentu C.

`[prelaz]`

---

## (16) Luka — prompt template — Luka (~3 min)

**Luka:**
Pogledajmo konkretno kako izgleda jedan prompt.

`[pokaži code block]`

Struktura ima šest delova. Prvo, *role priming* — „ti si iskusan instructional designer". Drugo, tema kursa i kontekst. Treće, lista ishoda učenja. Četvrto, **eksplicitan Bloomov nivo** za koji se generišu aktivnosti — „nivo: Analyze". Peto, *state-of-generation* — lista već generisanih aktivnosti sa uputstvom da se ne ponavljaju. Šesto, format izlaza — JSON sa zadatim poljima.

Zašto JSON, a ne markdown ili plain tekst? Tri razloga. Prvi — LLM-ovi su znatno bolji u generisanju validnog JSON-a kad im se eksplicitno traži, naročito modeli kao Qwen 2.5 i GPT-4. Drugi — JSON se trivijalno parsira i validira. Treći — JSON struktura se trivijalno mapira u OWL — to ćemo videti u Segmentu B.

`[pauza]`

Jedno opažanje iz iteracije. **State-of-generation context** je *najveća pojedinačna optimizacija* koju sam uveo u sistem. Pre nego što sam dodao tu jednu sekciju u prompt, sistem je generisao pet varijanti iste aktivnosti za isti ishod. Posle — koherentne, raznolike aktivnosti. Jedan tehnički trik koji menja kvalitet drastično.

`[prelaz na Uroša]`

---

## (17) Uroš — SOLO Quiz Generator i PS4 — Uroš (~4 min)

**Uroš:**
Sad ja, moj projekat.

`[pokaži opis levo, PS4 komponente desno]`

SOLO Quiz Generator. **Ulaz**: PDF lekcija. Sistem auto-detektuje da li je tekst na srpskom — latinicom ili ćirilicom — ili na engleskom, i prilagodi se. Parsira PDF u sekcije i learning objekte. Generiše ontologiju koncepata iz lekcije. **Izlaz**: MCQ pitanja na četiri SOLO nivoa, sa source_line citatima, sa ontology anchor-ima.

Tech stack — namerno **lokalni**: Ollama, Qwen 2.5 14B-instruct, kvantizovan na 4 bita da bi stao na običan GPU. Razlog je trostruk. Prvi — privatnost. PDF nastavnih materijala često sadrži IP fakulteta, ili materijal koji autor nije dao za cloud upload. Drugi — zero per-token cost. Treći — reproducibility. Lokalni model se ne menja preko noći kao GPT-4.

Ali — i to je važno reći — lokalni model je *slabiji* na nuansiranim zadacima. Što ćemo videti u evaluaciji.

`[pauza]`

Sad — *kako* generišem pitanja. Tu je središnja stvar — **PS4 prompt template.** Skraćeno od "*Persona, Pravila, Putanja, Pitanje*" — moj naziv, inspirisan praksom iz prompt engineering literature.

**Komponenta jedan — Persona.** Role priming. LLM-u se eksplicitno kaže: „ti si iskusan profesor, koji generiše pitanja za studente master nivoa, koja proveravaju određeni SOLO nivo." Nije ovo dekoracija. Promenjuje distribuciju izlaza merljivo.

**Komponenta dva — Pravila.** Kratka definicija SOLO nivoa, plus jedan worked example. Bitno — worked example je iz **drugog domena**, ne iz domena lekcije. Konkretno, koristim fotosintezu kao primer. Razlog: ne želim da model uči *sadržaj*, želim da uči *strukturu*. Da mu dam primer iz operativnih sistema, mogao bi da iskopira primer umesto da generiše novo pitanje.

**Komponenta tri — Putanja.** Chain-of-thought scaffold. LLM se ne traži da odmah napiše pitanje. Prvo se traži da identifikuje ključni koncept, pa da izvede SOLO nivo, pa tek onda da generiše pitanje. Tri koraka, eksplicitno.

**Komponenta četiri — Pitanje.** Strogi izlazni JSON schema sa poljima: question, correct_answer, distractors (lista), source_line, tags, level. Plus — *typed distractor strategies*. Za svaki SOLO nivo, postoji tabela tipova grešaka koje distraktori treba da reflektuju. Za Unistructural — leksička konfuzija, slično ime drugog koncepta. Za Multistructural — nepotpuna lista. Za Relational — pogrešna kauzalnost. Za Extended Abstract — preoptšta generalizacija.

To je PS4. Čitljiv akronim, prilično standardna pedagoška ideja, ali — sklopljena u jedan prompt template.

`[prelaz]`

---

## (18) Two-pass EA — Uroš (~3 min)

**Uroš:**
Sad jedan trik za najteži nivo — Extended Abstract.

Generisanje Extended Abstract pitanja je *najteže*. Razlog: model treba istovremeno da napravi pitanje koje generalizuje, *i* tri uverljiva distraktora koji testiraju različite vrste pogrešnog generalizovanja. To je previše za jedan prolaz — distraktori ispadaju ili identični tačnom odgovoru ili trivijalno slabi.

Moje rešenje: **dva prolaza.**

`[pokaži dijagram Pass 1 → arrow → Pass 2]`

**Pass 1.** LLM dobija sekciju iz PDF-a, SOLO definiciju, worked example. Vraća: pitanje, tačan odgovor, source_line. **Bez distraktora.** Fokus je samo na valjanosti pitanja i ankorovanju u izvor.

**Pass 2.** LLM dobija svoje vlastito izlazno pitanje iz prošlog prolaza, plus tabelu *typed distractor strategies*. Vraća: tri distraktora, svaki sa eksplicitnim tipom greške.

To je *predictive prompting* — model prvo zamišlja gde će student da pogreši, pa tek onda piše distraktore koji testiraju te tipove grešaka.

`[pauza]`

Rezultat: distraktori postaju merljivo bolji. U pilot evaluaciji, *uverljivost distraktora* — to je metrika gde nezavisni recenzent ocenjuje koliko su distraktori plauzibilni — skočila je sa prosečnih 2.8 na pet, na 3.9 na pet. To je velika razlika kad pričamo o ravnoj površini između „distraktor je očigledno pogrešan" i „distraktor zaista zbunjuje studenta koji nije razumeo dubinski."

`[prelaz]`

---

## (19) Source-line citation — Uroš (~3 min)

**Uroš:**
Najvažniji anti-halucinacioni mehanizam u mom projektu — i možda najjednostavniji.

**Source-line citation.** Za svako pitanje koje LLM proizvede, mora da vrati *doslovni navod iz PDF-a* koji opravdava tačan odgovor.

`[pokaži primer JSON levo, opis desno]`

Pogledajte primer. Pitanje: „Kako se memorijska hijerarhija odnosi prema brzini pristupa i kapacitetu?" Tačan odgovor: „Brzina opada, a kapacitet raste niz hijerarhiju." Source_line: „Memorijska hijerarhija je organizovana tako da brže memorije imaju manji kapacitet." — to je doslovni navod iz lekcije.

`[pauza]`

Zašto je ovo važno? Tri razloga.

**Prvo** — ako navod ne postoji u izvoru, **to je signal halucinacije.** Pravim *fuzzy match* između source_line i originalnog PDF teksta. Ako se ne nađe, pitanje se markira kao sumnjivo. Nije savršeno — ima false positive-a kad LLM parafrazira — ali je *dovoljno dobro* da uhvati prave halucinacije, kad LLM izmisli činjenicu.

**Drugo** — recenzent vidi *gde je istina*. Kad nastavnik gleda generisana pitanja, ne mora da pamti šta je bilo u PDF-u. Vidi pitanje, vidi citat. Verifikacija ide u sekundama, ne u minutama.

**Treće** — studentu se može pokazati objašnjenje koje je vezano za originalni izvor. Kad student promaši, kviz mu može pokazati ne samo „tačan odgovor je X", nego i „evo gde to piše u materijalu, sekcija 4.3, stranica 67."

To je deo šire ideje — **traceability**. Sve što sistem proizvede mora da se može pratiti unazad do nekog izvora — bilo PDF-a, bilo ontološke relacije, bilo eksplicitne pravila.

`[prelaz na Stefana]`

---

## (20) Stefan — dekompozicija zadataka — Stefan (~4 min)

**Stefan:**
Sad ja. Moj projekat — Computer Use Video Instructions.

`[pokaži dekompoziciju desno]`

Ulaz je prirodno-jezička instrukcija. Konkretno, evo primera koji koristim za demonstracije: *„Kreiraj C# konzolni projekat u Visual Studio-u koji ispisuje Hello World."*

Izlaz je MP4 video tutorijal — uživo snimak ekrana dok sistem zaista otvara Visual Studio, klika kroz dijaloge, kuca kod, pokreće program. Plus OWL fajl izvršnog plana — koji je auditable, što ćemo videti u Segmentu B.

Između — **trostepena dekompozicija.**

**Task** — kompletan zadatak. Zatvorena celina. Ima ime, opis, listu koraka.

**Step** — atomska akcija. Jedan klik. Jedan kuc. Jedna pauza. Step ima eksplicitan redni broj, eksplicitnu akciju, eksplicitan target.

**Action** — tip akcije. Ne proizvoljan tekst. Bira se iz **fiksne ontologije akcija**: `open_application`, `click`, `type_text`, `key_press`, `wait`, `capture_screen`. Šest tipova. To je sve. LLM ne može da izmisli sedmu akciju.

`[pauza]`

Ovaj dizajn — sa ograničenim setom akcija — je svesna kontrola. **Sistem može da izvrši samo ono što razume.** Ako LLM kaže „pomeri prozor u gornji desni ugao", sistem to ne razume kao akciju i odbija plan. Bolje da odbije nego da pokuša i napravi pogrešno.

Hijerarhijska razgradnja takođe služi kao *interpretabilan log*. Kad nešto pukne, vidite tačno na kom Step-u — sa kojim Action-om, kojim target-om — sistem je zastao. To je vrednije od stack trace-a.

LLM koji koristim za ovo je *cloud* — Groq Llama 3.3 70B. Razlog: za dekompoziciju zadataka, treba mi *velik* model — 70B parametara. Lokalni Qwen 14B koji Uroš koristi nije dovoljno dobar na ovom konkretnom zadatku. Ovde nema PDF nastavnog materijala čija je privatnost na kocki — instrukcija je sama nastavnička, nije privatan podatak.

Vision model je drugi — OpenRouter Qwen 2.5 VL 72B — koji *vidi* ekran i nalazi UI elemente za klikove. To ćemo razmotriti u Segmentu C kad pričamo o robustnosti.

`[prelaz]`

---

## (21) Prompt vs Fine-tune — Uroš (~3 min)

**Uroš:**
Pre nego što zatvorimo segment, dva tradeoff slajda.

Prvi — **prompt engineering vs fine-tuning.** Zašto smo, sva trojica, izabrali prompt engineering, a ne fine-tune-ovali sopstveni model?

`[pokaži tabelu]`

**Trošak razvoja.** Prompt — samo pišete prompt. Fine-tune — treba vam dataset, treba vam GPU za trening, treba vam evaluation harness.

**Trošak izvršenja.** Prompt — što veći prompt, više tokena. Fine-tune — manji prompt, ali specijalizovan model koji košta da se hostuje.

**Brzina iteracije.** Prompt — sekunde. Promenite tekst, pokrenete iznova. Fine-tune — sati do dani po iteraciji.

**Domen-specifičnost.** Prompt — ograničena. Oslanja se na bazni model. Fine-tune — visoka, ali samo za domen koji ste pokrili treningom.

**Kontrola halucinacija.** Prompt — posredna. Kroz CoT, schema, grounding. Fine-tune — direktna. Kvalitet treninga.

`[pauza]`

Bitno — postoji i treći put. **RAG, Retrieval-Augmented Generation.** Mi ga *posredno radimo* kroz ontology anchor i source_line. Nije RAG u klasičnom smislu, ali postiže sličan efekat — generisanje koje je usidreno u eksternom izvoru.

Naš izbor — *prompt engineering plus ontology grounding* — je donet zbog jednostavnog razloga: **ne postoji stabilan domenski dataset za SOLO/Bloom u programiranju**. Da fine-tune-ujemo, prvo bismo morali da napravimo dataset od nekoliko hiljada pitanja anotiranih po SOLO nivoima u domenu programiranja. To nema. Pa bismo onda fine-tune-ovali, pa morali da održavamo model kako se baza znanja proširuje. Skupo i krhko.

Prompt engineering je *brza, transparentna, jeftina* iteracija. Za istraživački projekat — savršeno. Za produkciju u velikoj instituciji sa stabilnim curriculum-om — verovatno bi fine-tune isplatilo dugoročno.

`[prelaz]`

---

## (22) Lokalni vs cloud — Stefan (~3 min)

**Stefan:**
Drugi tradeoff — **lokalni vs cloud modeli.** Tu nas trojicu interesuje jedna konkretna stvar: mi *koristimo i jedan i drugi* — što je netipično.

`[pokaži tabelu]`

Uroš koristi lokalni — Ollama Qwen 2.5 14B. Za SOLO pitanja. Razlog: PDF nastavni materijal je privatan, treba mu visok kvalitet ali ne *najviši*, i bitna mu je reproducibility za istraživanje.

Ja koristim cloud — Groq Llama 3.3 70B za planiranje koraka, plus OpenRouter Qwen 2.5 VL 72B za vision. Razlog: planiranje zahteva *veći* model, a vision modeli koji bi mi trebali nisu dostupni lokalno za realistično vreme — Qwen 2.5 VL 72B na lokalu zahteva više od 80 GB VRAM-a.

Trošak: Uroš plaća zero per call — sve lokalno. Ja plaćam malo po pozivu na Groq-u — Llama 70B je ovde subvencionisan brzom hardverskom inferenecijom. Vision je sredinj po ceni — par centi po slici.

Brzina: Lokalni je sporiji — zavisi od GPU-a. Cloud Groq je *vrlo* brz, sub-sekundno za prosečne zahteve.

Privatnost: lokalni — visoka, sve na mašini. Cloud — niska, podaci idu na server.

Kvalitet: lokalni je solidan za Unistructural i Multistructural, slabiji za Extended Abstract. Cloud — veoma dobar, naročito za vision.

`[pauza]`

Naša pouka iz iteracije: **ne tražite jedan model za sve.** Različiti zadaci traže različite modele. Pipeline koji integriše dva ili tri različita LLM-a, sa različitim privatnost-kvalitet-cena profilima — to je realnija arhitektura za produkciju.

`[prelaz na ključnu poruku segmenta]`

---

## (23) Ključna poruka Segmenta A — Uroš (~2 min)

**Uroš:**
Da rezimiramo Segment A.

`[pokaži centralnu poruku]`

**Hijerarhijska dekompozicija plus formalna pedagoška taksonomija ugrađena u prompt transformišu LLM iz generičkog generatora teksta u pedagoški kalibrisan alat.**

Tri mini-poente:

**Bloom** drives the WHAT — koje tipove aktivnosti pravimo. Lukin projekat ga koristi za strukturu kursa.

**SOLO** drives the HOW DEEP — koliko duboko ih proveravamo. Moj projekat ga koristi za pitanja.

**PS4 i dekompozicija** drives the HOW — kako pišemo prompt. Svi koristimo neku verziju ovoga — ja PS4 za pitanja, Luka state-of-generation za strukturu, Stefan trostepenu dekompoziciju za video.

`[pauza]`

Ali — i to je važno — **ovo nije dovoljno samo po sebi.** LLM, čak i sa savršenim promptom, može da haluciniše. Može da generiše izlaz koji je strukturalno tačan ali sadržajno besmislen. To nas vodi u Segment B — gde dodajemo *formalni grounding*.

`[prelaz na Segment B]`

---

# SEGMENT B — ONTOLOGIJE I FORMALNI GROUNDING

## (24) Divider Segment B — Stefan (~30 s)

**Stefan:**
Segment B. Ontologije i formalni grounding. Glavna teza: **ontologija funkcioniše kao ugovor između LLM-a i sistema.** Ono što LLM proizvede mora da se mapira na klase i relacije iz ontologije. Što filtrira halucinacije i omogućava ponovnu upotrebu sadržaja.

`[prelaz]`

---

## (25) OWL — Stefan (~3 min)

**Stefan:**
Brzi uvod u OWL — Web Ontology Language — taman koliko publika treba da prati priču.

OWL je W3C standard iz 2004, prošao kroz OWL 2 reviziju 2009. Definiše:

**Klase** — kategorije entiteta. U mom projektu: *Task*, *Step*, *Action*, *UIElement*, *Application*, *ExecutionState*.

**Instance** — konkretni entiteti tih klasa. *Step_1*, koji je tip *Step*. *open_application_Action*, koji je tip *Action*.

**Object properties** — relacije između entiteta. *hasStep*, koja povezuje *Task* sa *Step*. *hasAction*, koja povezuje *Step* sa *Action*.

**Data properties** — atributi entiteta. *stepOrder*, koji je integer. *targetName*, koji je string.

**Ograničenja** — pravila. *Task ima bar jedan Step*. *Svaki Step ima tačno jedan stepOrder*.

`[pauza]`

OWL je serijalizovan u nekoliko formata. Mi koristimo dva:

**Turtle** — čitljiv za čoveka, koristim ga za seed ontologiju. U primeru na ekranu vidite mali isečak iz `computer_use.ttl`.

**RDF/XML** — XML-zasnovan, generiše ga sistem za eksport. Otvara se u Protégé editoru.

Bitna karakteristika OWL-a — **reasoning.** Postoje reasoner-i — kao Pellet, HermiT — koji mogu da izvode *nove* činjenice iz postojećih, primenom pravila ontologije. Mi to *posredno* koristimo kroz SPARQL upite — ne pokrećemo reasoner eksplicitno.

`[prelaz]`

---

## (26) RDF i SPARQL — Stefan (~3 min)

**Stefan:**
RDF i SPARQL — kratko.

**RDF — Resource Description Framework.** Sve činjenice su **tripleti**: subjekat, predikat, objekat.

`[pokaži primer tripleta]`

*RAM, isFasterThan, Disk.* Tri elementa. To je sve. Bilo koja informacija u svetu može da se izrazi kao skup ovakvih tripleta. Knowledge graph od milijardu tripleta je jedan veliki RDF graf.

Razlika OWL i RDF — OWL je *TBox*, šema. RDF je *ABox*, instance. Klase vs konkretni entiteti. Ali oba se serijalizuju kao tripleti.

`[pauza]`

**SPARQL — SPARQL Protocol and RDF Query Language.** Upitni jezik za RDF grafove. Sintaksa liči na SQL, ali radi nad grafom, ne tabelom.

`[pokaži SPARQL primer]`

Pogledajte primer. *PREFIX cu* — definicija namespace-a. *SELECT ?step ?order ?action ?target* — koje varijable hoću. *WHERE blok* — uslovi koje pattern u grafu mora da zadovolji.

U ovom konkretnom primeru, upit kaže: nađi sve Step-ove koji pripadaju nekom Task-u, sa svojim redom, akcijom, i target-om. Sortiraj po redu.

Ovo je tačan upit koji moj sistem izvršava da bi pročitao plan iz OWL fajla i izvršio ga korak po korak.

SPARQL je *deklarativan* — opisujete šta hoćete, ne kako. Reasoner se brine o pretrazi. To je elegantno za naš slučaj — ne moramo da pišemo logiku obilaska grafa.

`[prelaz na Luku]`

---

## (27) Luka — JSON serijalizacija — Luka (~3 min)

**Luka:**
Sad ja, da pokažem kako moja struktura kursa izgleda u JSON formatu.

`[pokaži JSON]`

Hijerarhijska struktura. Spolja je *course* — sa nazivom. Unutar njega lista *modules*. Svaki modul ima naziv i listu *concepts*. Svaki koncept ima naziv i listu *outcomes*. Svaki ishod ima naziv, Bloomov nivo, i listu *activities*. Aktivnosti su listovi u stablu.

Postoji *jedno* pitanje koje ovde uvek dođe — *zašto JSON, a ne odmah OWL*?

Tri razloga.

**Prvo** — LLM-ovi su znatno bolji u generisanju validnog JSON-a nego validnog OWL-a. JSON je *u podacima na kojima su trenirani.* OWL — manje. Generisanje JSON-a je niži-rizik operacija.

**Drugo** — JSON se trivijalno prikazuje u UI. React komponente, drvasti pregled, editor. OWL bi tu bio glomazan.

**Treće** — *mapiranje JSON → OWL je deterministički*. Postoji jedan-na-jedan mapiranje između JSON čvorova i OWL klasa. Pišete jednu mapping funkciju, jednom, i imate OWL kad god vam treba. Modul postaje *Module* klasa, koncept postaje *Concept* klasa, *hasModule* property povezuje Course i Module... i tako dalje.

Dakle: LLM generiše JSON, mapper pravi OWL. Najbolje od oba.

`[prelaz]`

---

## (28) Uroš — ConceptRelationship ontologija — Uroš (~4 min)

**Uroš:**
Sad ja, da pokažem moju ontologiju.

`[pokaži levo opis, desno tipove relacija]`

Moja ontologija nije unapred fiksna — *gradi se iz materijala.* Sistem uzme PDF lekciju, ekstraktuje *learning objekte* — atomske jedinice znanja — i onda ekstraktuje **relacije između njih.**

Svaka relacija je *edge* u grafu. Ima izvor, cilj, tip, i opis. Čuva se u SQL tabeli `ConceptRelationship`, i može se eksportovati u Turtle ili OWL.

Tipovi relacija:

**prerequisite** — A se mora razumeti pre B. Najčešća relacija.

**example_of** — A je primer pojma B. „PCB je primer kontrolne strukture."

**depends_on** — A operativno zavisi od B. „Context switch zavisi od PCB-a."

**similar_to** — A i B dele ključnu karakteristiku. „Semaphore similar_to Mutex."

**part_of** — A je deo strukture B. „PID part_of PCB."

`[pauza]`

Ekstrakcija ide u *pet prolaza* — jedan po tipu relacije. LLM se eksplicitno traži: „pronađi sve prerequisite veze između ovih koncepata." Pa onda — „pronađi sve example_of veze." I tako dalje.

Razlog za multi-pass: kad LLM-u dam zadatak da nađe sve relacije odjednom, *fabricira* prerequisite lance — kreira lažne hijerarhije gde ih nema, jer je trening data full of takvih lanaca. Multi-pass pristup ga drži *fokusiranim na jedan tip* po prolazu — daje mnogo bolje rezultate.

Postoji i konzervativni fallback: ako LLM vrati nulu za neki tip, sistem ne fabrikuje veze — koristi samo statičke heuristike, na primer „dva learning objekta sa istim ključnim rečima imaju shared_keyword vezu." Bolje malo veza nego pogrešne veze.

Rezultat je *knowledge graph specifičan za jednu lekciju* — koji se onda koristi kao kontekst za pitanja.

`[prelaz]`

---

## (29) Ontology anchor — Uroš (~3 min)

**Uroš:**
Kako se ontologija koristi za pitanja?

`[pokaži SPARQL upit]`

Za **relacionalna** i **Extended Abstract** pitanja, sistem prvo izvršava SPARQL upit nad lekcijinom ontologijom. Upit izgleda otprilike ovako: *nađi jednu nasumičnu ConceptRelationship koja pripada ovoj lekciji.*

To što upit vrati — recimo, *(RAM, isFasterThan, Cache, „RAM ima brži pristup nego cache memorija")* — postaje **ontology anchor.**

Anchor se ubacuje u prompt LLM-a sa instrukcijom: „napravi relacionalno pitanje koje testira *baš ovu vezu.*"

`[pauza]`

Posledice — tri.

**Prvo, traceability.** Svako relacionalno pitanje koje sistem proizvede *zna se na koju vezu se odnosi.* Anchor se čuva u `tags.ontology_anchor` polju pitanja. Nastavnik kliknete na pitanje i vidi: ovo pitanje je vezano za relaciju RAM isFasterThan Cache. Auditable lanac.

**Drugo, dedup.** Kako da znate da li su dva pitanja duplikati, čak i kad su tekstualno različita? Moj prvi pokušaj je bio *word overlap* — broji koliko reči se preklapa. Loše. Tačan dedup je: **(ontology_anchor, normalized_correct_answer)** ključ. Dva pitanja sa istim anchor-om i istim normalizovanim tačnim odgovorom su duplikati, bez obzira na formulisanje.

**Treće, recenzija nastavnika je drastično lakša.** Umesto da se nastavnik probija kroz tekst, on vidi: „ovo pitanje testira sledeću vezu — RAM brži od Cache. Citat iz materijala: ovo. Tip distraktora: kauzalna inverzija." Pogled na ekran, klik gore-dole, done.

`[prelaz na Stefana]`

---

## (30) Stefan — OWL plana — Stefan (~4 min)

**Stefan:**
Sad ja, da pokažem moju OWL ontologiju.

`[pokaži levo opis, desno OWL primer]`

Moja ontologija je *fiksna* — `computer_use.ttl`. Pisana je rukom, ne generisana iz materijala. Razlog: domen je *zatvoren* — automatizacija desktop interakcija ima konačan vokabular akcija. Ne treba mi dinamička ontologija kao Urošu.

Klase:

**Task** — kompletan automatizacioni zadatak. „Kreiraj C# konzolni projekat u Visual Studio-u."

**Step** — atomska akcija u zadatku. „Otvori Visual Studio." „Klikni na 'Create new project'."

**Action** — tip akcije iz fiksne liste. Već smo videli šest tipova.

**UIElement** — UI element. Dugme, polje, meni. Identifikuje se imenom i opcionalno koordinatama.

**Application** — aplikacija. Visual Studio, Chrome, Notepad, ... Step ima `requiresApplication` property.

**ExecutionState** — stanje izvršenja koraka. *Pending*, *Executing*, *Completed*, *Failed*. Ovo je važno za auditability — nakon izvršenja, OWL fajl se *prepisuje* sa novim stanjima, i sačuva se kao `*_executed.owl`.

`[pauza]`

Object properties — *hasStep*, *hasAction*, *hasTarget*, *nextStep*, *requiresApplication*.

Data properties — *stepOrder*, *stepDescription*, *targetName*, *inputValue*, *waitDuration*, *expectedResult*.

Primer instance Step-a sa svim ovim — vidite na desnoj strani slajda. To je tačan deo `.ttl` fajla koji sistem generiše za prvi korak iz primera „Kreiraj C# projekat..."

`[prelaz]`

---

## (31) Validacija — Stefan (~3 min)

**Stefan:**
Pravila validacije. Pre nego što sistem išta izvrši, plan se *validira* SPARQL upitima nad ontologijom.

`[pokaži tabelu pravila]`

**Action validation.** Da li su sve akcije iz fiksne liste? Ako LLM kaže „swipe_screen", sistem to ne razume. Stop, ne radi se.

**Sequence validation.** Da li je `open_application` *pre* svake interakcije sa tom aplikacijom? Ne možete da kliknete na meni u Visual Studio-u ako Visual Studio nije otvoren. SPARQL upit traži *prvu interakciju* sa svakom aplikacijom, i proverava da li je pre nje bio open_application.

**Wait recommendation.** Nakon `open_application` preporučuje se `wait` od bar dve sekunde. Razlog: UI se učitava — bez pauze, klikovi promašuju. Ako LLM nije ubacio wait, sistem ubacuje implicitno.

**Parameter validation.** `type_text` zahteva `inputValue`. `key_press` zahteva ime tipke. `click` zahteva `targetName`. Bez parametra, akcija ne može da se izvrši — sistem odbija plan.

**Target validation.** `click` mora imati specifičan target. Ne može „klikni negde gore." Mora biti element koji vision model može da nađe.

`[pauza]`

Sve ovo je *pre izvršenja*. Ne lečimo neuspeh nakon što se desi — *sprečavamo* ga pre nego što se desi.

To je suština onoga što sam nazvao **ontologija kao ugovor.** LLM može da kaže šta hoće — ali ako to ne mapira na ovu ontologiju, sistem to ne izvršava. Halucinacija filtrirana pre nego što je stigla na ekran.

`[prelaz]`

---

## (32) Anti-halucinacije: tri ugla — Uroš (~3 min)

**Uroš:**
Sad nešto što je možda *centralna poenta cele prezentacije*. Svako od nas se bori protiv halucinacija, ali *na drugi način*. Tri komplementarna pristupa.

`[pokaži tri kolone]`

**Ja — source_line citat.** Već smo videli. LLM mora vratiti doslovni navod iz PDF-a. Ako navod ne postoji u izvoru, sistem markira pitanje. Adresira: **faktičke halucinacije** — kad LLM izmisli činjenicu.

**Luka — state-of-generation kontekst.** Prethodno generisane aktivnosti se ubacuju u prompt sa eksplicitnim „ne ponavljaj." Adresira: **redundanciju i nekoherentnost** — kad LLM proizvede pet varijanti istog umesto da pokrije skup različitih nivoa.

**Stefan — SPARQL validacija.** Plan se proverava pre izvršenja. Adresira: **strukturne halucinacije** — kad LLM proizvede tehnički validan tekst ali sekvenca akcija nema smisla.

`[pauza]`

Ovo nije slučajno. Tri tipa halucinacije zahtevaju *različite* odgovore.

Faktičku halucinaciju ne hvata SPARQL validacija — sintaktički je validna, samo je netačna. Hvata je samo *poređenje sa izvorom.*

Strukturnu halucinaciju ne hvata source_line — citat iz PDF-a vam ne kaže da li ćete kliknuti u praznom prozoru. Hvata je samo *formalna logička provera plana.*

Redundanciju ne hvata ni jedno ni drugo — hvata je samo *memorija o već generisanom.*

Dakle, ako sistem treba da bude robustan, mora da kombinuje *više mehanizama*. Jedan mehanizam adresira jedan tip greške. Pravi pipeline ima sva tri.

To je tema koja se vraća kroz celu prezentaciju — i opravdava zašto je ova kombinacija „više od zbira svojih delova."

`[prelaz]`

---

## (33) Ključna poruka Segmenta B — Stefan (~2 min)

**Stefan:**
Rezime Segmenta B.

`[pokaži centralnu poruku]`

**Ontologija je ugovor između LLM-a i sistema.** Ono što LLM proizvede mora da se mapira na klase i relacije. To filtrira halucinacije i omogućava ponovnu upotrebu sadržaja.

Tri mini-poente:

**OWL/RDF** — definicija strukture i instanci. *Šta su validni objekti u sistemu.*

**SPARQL** — kontrola toka i upita. *Kako se kreće kroz strukturu i kako se provera radi.*

**Anchor** — vezivanje pitanja za izvor. *Specifična instance koja ankoruje generisanje.*

`[pauza]`

Sad — sve je ovo lepa teorija. Pitanje koje moramo da postavimo je — **da li radi?** I koliko dobro? Šta su ograničenja? Tu nas vodi Segment C.

`[prelaz u Segment C]`

---

# SEGMENT C — EVALUACIJA, KVALITET I OGRANIČENJA

## (34) Divider Segment C — Uroš (~30 s)

**Uroš:**
Segment C. Evaluacija, kvalitet, ograničenja. Glavna teza: **kvalitet automatski generisanog sadržaja opada sa porastom kognitivne složenosti.** To nije bug, to je strukturna posledica — i upravo zato je human-in-the-loop obavezan.

`[prelaz]`

---

## (35) Pilot evaluacija — Uroš (~4 min)

**Uroš:**
Kako merimo kvalitet SOLO pitanja?

`[pokaži opis levo, tabela desno]`

**Cilj evaluacije** — dva pitanja. Prvo: da li je kvalitet konstantan kroz SOLO nivoe? Drugo: da li distraktori postaju lažni — to jest, trivijalno netačni ili odbranjivi kao tačni — na višim nivoima?

**Metodologija.** Više PDF lekcija iz različitih domena — operativni sistemi, baze podataka, algoritmi, mašinsko učenje. Sistem generiše N pitanja po nivou po lekciji. Ručna anotacija — sudije su članovi tima plus jedan nezavisni recenzent. Metrike u četiri dimenzije:

**Valjanost** — da li je pitanje semantički ispravno, da li meri ono što tvrdi da meri.

**Kvalitet distraktora** — koliko su uverljivi, koliko su pravilno tipizirani.

**Source-line tačnost** — da li je navod iz PDF-a koji LLM vraća zaista u PDF-u.

**SOLO usklađenost** — da li je pitanje na SOLO nivou koji tvrdi da je.

Svaka dimenzija — skala jedan do pet. Plus komentari za kvalitativnu analizu.

Format izlaza — tabela sa pitanjima u redovima i metrikama u kolonama, plus komentari. Pridružujemo aggregirane statistike po nivoima i po domenima.

`[pauza]`

Bitno reći: **ovo nije fakultetska studija sa stotinama studenata.** Ovo je *pilot evaluacija*, ručna anotacija, mali broj pitanja, mali broj sudija. Cilj nije generalizacija — cilj je *otkrivanje obrazaca grešaka* da bismo iterirali na sistemu. Real-world evaluacija u učionici ostaje *otvoreno pitanje za budući rad.*

`[prelaz]`

---

## (36) Rezultati po SOLO nivou — Uroš (~4 min)

**Uroš:**
Šta vidimo u rezultatima.

`[pokaži bar chart levo, opis desno]`

**Unistructural** — visok kvalitet. Oko devedeset posto. Sistem dobro generiše „setiti se jedne činjenice." Distraktori su uverljivi, source_line gotovo uvek tačan. Mali domen za promašaj.

**Multistructural** — i dalje solidno. Oko osamdeset pet posto. Glavna greška: distraktori su povremeno *podskupovi* tačnog odgovora — kao „tri stanja procesa: Ready, Running" — što je tačno ali nepotpuno.

**Relational** — *primetan pad.* Oko sedamdeset posto. Pitanja su valjana, ali distraktori počinju da budu odbranjivi kao tačni. „Kako kontekstno prebacivanje utiče na throughput?" — i distraktor *„nema uticaja, jer je trošak zanemarljiv"* može da bude tačan u određenim kontekstima.

**Extended Abstract** — *najveći pad.* Oko pedeset pet posto. Pitanja su često previše generička. Distraktori su retko uverljivi — ili su očigledno pogrešni, ili su skoro identični tačnom odgovoru. Two-pass pristup pomaže, ali ne rešava problem potpuno.

`[pauza]`

Ovo *nije iznenađujuće.* Lister i kolege su to opservirali u 2006 — Extended Abstract zadaci su strukturalno teški *za studente*, i sad vidimo da su strukturalno teški *i za LLM-ove*. Razlog je u suštini isti — generalizacija zahteva sposobnost preuzeti koncept iz jednog konteksta i primeniti ga u drugom. To je *najteže* za sve aktere, ljude i mašine.

**Posledica za sistem.** Na Unistructural i Multistructural nivou, generisana pitanja su gotovo direktno upotrebljiva. Na Relational — preporučuje se recenzija. Na Extended Abstract — **obavezna** recenzija nastavnika, i očekujte da ćete trećinu pitanja odbaciti.

To nije slabost sistema — to je *opravdana podela rada* sa nastavnikom. Sistem snima 80% mehanike. Najteže odluke ostaju čoveku.

`[prelaz]`

---

## (37) Primeri SOLO pitanja — Uroš (~3 min)

**Uroš:**
Da vidimo konkretne primere, iz domena operativnih sistema.

`[pokaži 4 kartice sa Q i A]`

**Unistructural.**
Pitanje: *„Koja struktura procesa čuva PID?"*
Tačan odgovor: *„PCB — Process Control Block."*
Source_line: „PID je smešten u Process Control Block, zajedno sa ostalim atributima procesa."
Distraktori: TCB, BIOS, MMU. Sva tri su sintaktički slični, ali ni jedan nije PCB.
*Komentar:* dobro pitanje, čisti recall. Distraktori su tipizirani — leksička konfuzija sa drugim trobokvenim akronimima.

**Multistructural.**
Pitanje: *„Koja tri stanja procesa najčešće prepoznajemo?"*
Tačan odgovor: *„Ready, Running, Blocked."*
Distraktor 1: „Ready, Running, Sleeping" — *terminologijska konfuzija* (Sleeping je donja klasa Blocked-a).
Distraktor 2: „New, Running, Terminated" — *nepotpuna lista* (preskačemo Ready i Blocked).
Distraktor 3: „Active, Inactive, Suspended" — *pogrešan model* (Windows process states vs UNIX).
*Komentar:* svaki distraktor reflektuje *različitu vrstu greške* — to je upravo ono što typed distractor strategije postižu.

**Relational.**
Pitanje: *„Kako kontekstno prebacivanje utiče na throughput sistema?"*
Tačan odgovor: *„Povećava overhead, smanjuje koristan rad CPU-a."*
*Ontology anchor:* (ContextSwitch, decreases, Throughput, „kontekstno prebacivanje smanjuje propustni opseg sistema kroz overhead")
*Komentar:* relacija je eksplicitna, pitanje proverava razumevanje uzročno-posledične veze.

**Extended Abstract.**
Pitanje: *„Kako biste prilagodili scheduling algoritam za real-time sistem sa zagarantovanim deadline-om?"*
Tačan odgovor: *„Korišćenje algoritama kao EDF ili Rate Monotonic, prioritetizovanih po roku, sa preemption-om."*
*Komentar:* pitanje zahteva primenu koncepta scheduling-a na *novi kontekst*. Distraktori su teški jer i pogrešni odgovori — kao „Round Robin sa kraćim quantum-om" — imaju svoju logiku. Tipičan kandidat za nastavničku recenziju.

`[prelaz na Luku]`

---

## (38) Luka — Bloom pokrivenost — Luka (~3 min)

**Luka:**
Sad ja, da pokažem evaluaciju moje strane.

`[pokaži opis levo, bar chart desno]`

**Metrika.** Da li su zastupljeni svi Bloomovi nivoi u generisanim kursevima? Koliko aktivnosti po nivou? Da li je sistem proizveo balansiranu strukturu, ili je „klizio" prema lakšim nivoima?

**Pristup.** LLM se *eksplicitno traži* za svaki nivo posebno. Sistem ne prelazi na sledeći nivo dok prethodni nije mapiran. Plus dedup po (ishod, aktivnost_tip).

**Nalaz.** Najbolji rezultati: Pamtiti i Razumeti. Oko dvanaest do petnaest aktivnosti po modulu, što je previše ako se uzmu sve — ali to je dobro, nastavnik bira podskup.

Primeniti — *takođe dobro*, oko osam do deset po modulu. Konkretne vežbe, problemi.

Analizirati — *primetan pad u kvalitetu, ne broju.* Sistem proizvodi aktivnosti tipa „uporedi X i Y" — što je formalno Analyze, ali često površno. Bez ozbiljnog mehanizma analize, samo paralelno pominjanje.

Evaluirati i Kreirati — *najslabije.* Aktivnosti tipa „kritički razmisli" ili „napravi sopstveno rešenje" su retke, i kad ih ima, generičke su. Tu je *limit prompt engineering-a* — viši Bloomovi nivoi zahtevaju zaista kreativnu sintezu, što LLM-ovi danas teško rade pri jednom pozivu.

`[pauza]`

Slično kao Uroševa SOLO evaluacija — sistem je *jak na nižim nivoima, slab na višim.* Razlog je strukturno sličan: viši nivoi zahtevaju sintezu i transfer, što su najteže operacije za current generation LLM-ova.

`[prelaz na Stefana]`

---

## (39) Stefan — robustnost UI — Stefan (~3 min)

**Stefan:**
Moja evaluacija ide drugim putem — nije o pedagoškom kvalitetu, nego o **robustnosti izvršenja.**

Vision model — Qwen 2.5 VL 72B preko OpenRouter-a — radi sledeće: dajem mu screenshot ekrana plus opis traženog UI elementa („dugme 'Create new project' u Visual Studio start prozoru"), on vraća bounding box. Onda PyAutoGUI klikne u centar tog boxa.

`[pokaži levo opis, desno tabelu testiranih aplikacija]`

Ovo radi *odlično* za standardne aplikacije.

**Visual Studio 2022** — stabilno. Većina dijaloga, menija, dugmadi prepoznata pouzdano.

**VS Code** — stabilno. Slično.

**Browseri — Chrome, Firefox, Opera, Edge** — stabilno za web case-ove. Adresna traka, dugmad navigacije, generička UI dugmad.

**Notepad i Notepad++** — stabilno. Minimalan UI, lako za vision model.

**Nestabilno — custom WinForms.** Aplikacije sa nestandardnim kontrolama — često se vide kao „područja" bez prepoznatljive strukture. Vision model promaši.

**Nestabilno — web aplikacije sa lazy UI.** Single-page aplikacije gde se sadržaj učitava asinhrono. Screenshot je uhvatio prelaznu stranicu, click promaši.

`[pauza]`

To je u skladu sa benchmarkovima koje smo *gledali kao referencu*. OSWorld benchmark, Xie i kolege 2024 — najbolji vision agenti danas postižu *oko dvadeset procenata success rate-a* na realnim desktop zadacima u open-ended setting-u. Anthropic Computer Use, oktobar 2024, dao je oko 14.9% — najnoviji Claude 3.5 Sonnet je u 20-30% rasponu.

Naš sistem nije agent u tom smislu — mi pravimo *eksplicitan plan* i izvršavamo ga. Ali isto ograničenje važi: vision model nije perfektan, naročito kod nestandardnih UI-jeva. Što treba transparentno reći.

**Posledica.** Kad sistem ne uspe da locira UI element, *zaustavlja se* — ne pokušava nasumično. Video se završava na tački prekida. Sistem loguje gde je tačno puklo. To olakšava debug i daje nastavniku jasan signal — „ovde vam treba ručna verzija ovog tutorijala."

`[prelaz na Uroša]`

---

## (40) PDF Coverage Tracking — Uroš (~3 min)

**Uroš:**
Još jedan komad evaluacije — *meta-evaluacija.*

`[pokaži opis levo, slika heatmap desno]`

**Šta meri.** Koje su stranice PDF-a pokrivene generisanim pitanjima? Koje nisu? Sistem to izračunava automatski.

Tri metrike:

**Pages covered** — sirov broj stranica koje je dotaklo bar jedno pitanje.

**Weighted coverage** — pokrivenost ponderisana brojem znakova po stranici. Naslovna sa pedeset znakova ne računa se isto kao sadržajna stranica sa dve hiljade.

**Substantive coverage** — ponderisana pokrivenost koja isključuje skoro prazne stranice — naslove, prazne stranice, slike bez teksta.

**Vizualno** — heatmap. Po stranici jedna traka. Visina trake = broj znakova na stranici. Boja = pokriveno ili nepokriveno.

Plus — lista *suštinskih stranica bez pitanja.* Ako imate stranicu sa dve hiljade znakova teksta a nijedno pitanje je ne pominje, sistem vam to direktno kaže — „stranica 47 ima 1840 znakova, nije pokrivena."

`[pauza]`

**Posledica.** Sistem vam *sam* signalizira gde da generišete dodatna pitanja. Nastavnik ne mora da gleda PDF i mentalno proverava — heatmap mu pokazuje.

Ovo je primer onoga što se često zove „**second pair of eyes**" — sistem ne samo da generiše, već i *kritikuje vlastiti izlaz.* To je dimenzija koja često nedostaje u LLM-driven sistemima — meta-procena pokrivenosti je *jeftina*, *brza*, i *vredna*.

Kad smo prvi put pokrenuli na realnoj lekciji, otkrili smo da sistem pokriva samo oko šezdeset posto stranica, iako generiše dvadeset pitanja. To je bilo iznenađujuće — pokazalo je da LLM ima tendenciju da se „fokusira" na određene delove materijala i ignoriše ostalo. Coverage tracking nam je dao alat da to *izmerimo* i *kompenzujemo.*

`[prelaz]`

---

## (41) Granice automatizacije — Stefan (~3 min)

**Stefan:**
Sad — pošteno priznanje. Šta sistem *ne može* da uradi.

`[pokaži levo listu, desno HITL flow]`

**Definisanje ishoda učenja.** Ovo zahteva pedagošku procenu — šta studenti treba da znaju nakon kursa? To je domen u kojem nastavnik ima ekspertizu koju LLM nema. Sistem može da pomogne — predloži ishode na osnovu materijala — ali konačnu odluku donosi čovek.

**Procena težine pitanja u kontekstu konkretne grupe.** Pitanje koje je „prelagano" za jednu generaciju može da bude „pretežo" za drugu. Sistem ne zna istoriju vaše grupe, vaš nastavni stil, vaš prethodni semestar.

**Procena Extended Abstract pitanja.** Već smo videli — kvalitet pada. Na ovom nivou nastavnik mora da pregleda svako pitanje.

**Recenzija video tutorijala.** Sistem može da snimi video, ali ne može da odgovori — *da li je didaktički dobar*? Da li je tempo dobar? Da li je redosled koraka pedagoški smislen? To je nastavnikova procena.

**Konačno korigovanje halucinacija.** Sistem može da hvata mnoge halucinacije — kroz source_line, ontology validation. Ali ne sve. Nastavnik je poslednja linija odbrane.

`[pauza]`

**Sve ovo radi čovek.** Sistem snima do osamdeset procenata mehaničkog rada. Čovek validira, koriguje, donosi stručne odluke.

Ovo je tema u literaturi pod nazivom **human-in-the-loop** — koja je eksplicitno na listi tema predmeta. Ne *human out of the loop* — to je opasna utopija. *Human in the loop* — sistem i čovek u sinergiji.

Konkretan flow:

**Korak jedan.** Sistem generiše — strukturu, pitanja, video.

**Korak dva.** Nastavnik pregleda — kroz source_line za pitanja, kroz ontology anchor chip, kroz coverage heatmap.

**Korak tri.** Nastavnik koriguje — edituje pitanja, prepravlja video instrukcije, dorađuje strukturu.

**Korak četiri.** Sistem beleži promene — u bazu, kao feedback signal. Što omogućava buduće generacije da uče od korekcija (još nije implementirano kod nas — to je deo budućeg rada).

`[prelaz]`

---

## (42) Ključna poruka Segmenta C — Uroš (~2 min)

**Uroš:**
Rezime Segmenta C.

`[pokaži centralnu poruku]`

**Kvalitet automatski generisanog sadržaja opada sa porastom kognitivne složenosti. To nije bug, to je strukturna posledica — i upravo zato je human-in-the-loop obavezan.**

Ova rečenica je *centralna* za diskusiju o AI u obrazovanju uopšte. Nije pitanje *da li* će čovek biti u petlji — pitanje je *gde.* Naš pipeline daje konkretan odgovor: *čovek je u petlji na visokim kognitivnim nivoima* — Extended Abstract pitanja, kreativna sinteza, pedagoška procena.

Na nižim nivoima — recall, prepoznavanje, čak i osnovna primena — sistem može da preuzme veliki deo posla. Što oslobađa nastavnika da se *fokusira* na delove gde njegova ekspertiza najviše vredi.

To nije pretnja struci nastavnika — to je redistribucija. *AI ne menja nastavnika, AI menja šta nastavnik radi.*

`[prelaz u Segment D]`

---

# SEGMENT D — SINTEZA, BUDUĆI RAD, ZAKLJUČAK

## (43) Divider Segment D — Luka (~30 s)

**Luka:**
Segment D. Sinteza, budući rad, zaključak. Kako bi tri komponente bile integrisane u jedinstven sistem, sa predlogom konkretne arhitekture. Veze sa drugim temama predmeta. Otvorena pitanja. I zaključak — vraćanje na glavnu poruku.

`[prelaz]`

---

## (44) Integrisana arhitektura — Luka (~4 min)

**Luka:**
Šta bi izgledalo da spojimo tri projekta u jedan sistem? Predlažemo **četvoroslojnu arhitekturu.**

`[pokaži levo opis, desno dijagram]`

**Sloj 1 — Course Planner.** Moj deo. Strukturu kursa generiše iz ishoda učenja i PDF materijala. Izlaz: JSON i OWL hijerarhija aktivnosti. Ovo je *ulazni* sloj — sve počinje ovde.

**Sloj 2 — Knowledge Layer.** ConceptRelationship ontologija iz lekcija. Urošev deo, generalizovan. Source-of-truth za relacije pojmova — koji koncept je preduslov kojem, koji je primer kojeg, šta zavisi od čega. Ovaj sloj *konsumiraju* svi naredni slojevi.

**Sloj 3 — Assessment Layer.** SOLO pitanja, ankored na knowledge layer. Moj kviz generator. Generiše pitanja za svaki ishod učenja koji Course Planner proizvede.

**Sloj 4 — Demonstration Layer.** Stefanov video pipeline. Praktične aktivnosti — one koje Course Planner identifikuje kao „praktična vežba" — postaju NL instrukcije, pa OWL plan, pa video tutorijal.

`[pauza]`

Bitno reći: **ovo nije implementirano.** Naša tri projekta su trenutno *odvojeni repoziti* — sa kompatibilnim, ali ne deljenim podacima. Integracija je *predlog* — naredni korak.

Ali — *zašto* mislimo da bi ovo radilo? Tri razloga.

**Prvi** — ontologija je *zajednički jezik.* Sva tri projekta koriste OWL/RDF, sa preklapajućim klasama (Concept, LearningObject). Mapping je trivijalan.

**Drugi** — taksonomije su *komplementarne.* Bloom u Course Planner-u, SOLO u Assessment-u, fiksna ontologija akcija u Demonstration-u. Ne sukobljavaju se — pokrivaju različite dimenzije.

**Treći** — *human-in-the-loop touchpoint-ovi* su isti. Nastavnik validira na istom mestu — review, edit, save. Ne radimo dva interfejsa.

`[prelaz]`

---

## (45) Tok podataka — Luka (~3 min)

**Luka:**
Kako bi tok podataka izgledao kroz integrisan sistem?

`[pokaži 6 koraka]`

**Korak jedan.** Nastavnik učitava materijal — PDF lekcije, listu ishoda, opis kursa.

**Korak dva.** Course Planner gradi strukturu — moduli, koncepti, ishodi, aktivnosti. Bloom-aware.

**Korak tri.** Knowledge Layer ekstraktuje ontologiju — multi-pass LLM ekstrakcija ConceptRelationship-a između learning objekata identifikovanih u Course Planner-u.

**Korak četiri.** Assessment generiše SOLO pitanja. Anchor iz Knowledge Layer-a, PS4 prompt, source_line.

**Korak pet.** Demonstration generiše video za praktičnu aktivnost. NL instrukcija (iz Course Planner-a, za aktivnosti tipa „praktična vežba") → OWL plan → SPARQL validacija → izvršenje + snimanje.

**Korak šest.** Nastavnik pregleda i koriguje. Coverage heatmap, source_line provera, video review.

`[pauza]`

Jedna važna stvar — *koraci 2 do 5 mogu da se paralelizuju.* Nije sekvencijalno. Coursе Planner može da završi modul i prosledi ga Knowledge Layer-u, dok parallelno radi sledeći modul. Background job pattern — *u skladu sa* arhitekturom koju Urošev SOLO Quiz Generator već koristi sa ThreadPoolExecutor-om.

`[prelaz]`

---

## (46) Veze sa drugim temama — Uroš (~4 min)

**Uroš:**
Sad nešto što direktno povezuje naš pipeline sa *drugim temama* sa liste predmeta.

`[pokaži dve kolone]`

**Deep Reinforcement Learning za ITS-ove.** Sa liste je *Deep Reinforcement Learning for Intelligent Tutoring Systems.*

Naš pipeline trenutno *generiše* materijal — ali ne *bira* materijal *adaptivno za konkretnog studenta.* To je sledeći korak. Konkretno:

**State** — profil studenta plus istorija odgovora. Šta zna, šta ne zna, koje SOLO nivoe je savladao.

**Action** — izbor narednog pitanja ili aktivnosti. Iz bazena koji je naš sistem generisao.

**Reward** — napredak studenta na SOLO i Bloomovoj skali. Plus, sekundarni signal — vreme provedeno na zadatku, broj pokušaja.

Ovde DRL postaje *meta-sistem nad našim pipeline-om*. Naš sistem proizvodi *bazene mogućih sledećih koraka*, DRL bira *koji od njih* za konkretnog studenta.

`[pauza]`

**Recommender Systems** za personalizaciju putanja učenja. Drugi tema sa liste — *Recommenders for ITS.*

Nije DRL — recommender. Razlika je važna.

**Content-based.** Ontology anchor pitanja služi kao feature. „Student je promašio pet pitanja koja su ankoroovana za 'process_synchronization' koncept — preporuči mu materijal koji pokriva taj koncept."

**Collaborative.** Studenti sa sličnim odgovorima dele preporuke. „Studenti koji su promašili pitanja kao ti, popravili su se kroz aktivnost X."

**Hibridno.** Koristi oba signala — sadržaj i kolaboraciju.

`[pauza]`

I treća veza, koja nije eksplicitno na listi ali je relevantna — **Data-Centric AI.** Naš sistem proizvodi *strukturisane podatke* — pitanja sa anchor-ima, video sa OWL planovima. Ti podaci sami po sebi mogu da budu *trening set za sledeći model* — pristup koji Andrew Ng zove „data-centric AI." Ne menjamo algoritam, već poboljšavamo *podatke*.

Konkretno — Urošev sistem trenutno generiše pitanja sa source_line. Kada nastavnik koriguje pitanje, ta korekcija postaje *trening signal.* Time bismo mogli da posredno fine-tune-ujemo specifične promptove, bez fine-tune-ovanja celog modela.

`[prelaz]`

---

## (47) Otvorena pitanja — Stefan (~3 min)

**Stefan:**
Šta sve nismo rešili.

`[pokaži 5 pitanja]`

**Q1. Kvalitet distraktora na EA nivou.** Kako podići uverljivost bez gubitka SOLO tipa greške? Trenutno — two-pass plus typed strategies — pomaže, ali ostaje gap. Potencijalno rešenje: aktivno uključiti nastavnika u prompt — „nastavnice, koji su tipični nepravilni načini razmišljanja koje studenti pokazuju na ovoj temi?" — i koristiti to kao input.

**Q2. Real-world evaluacija.** Sistem nije testiran u stvarnoj učionici — što je sledeći korak. Konkretno, planiramo da postavimo pilot na predmetu Računarski sistemi ili sličnom — gde bismo nastavili paralelno: pola pitanja generisanih sistemom, pola ručno, i merili performans studenata.

**Q3. Multimodalna procena znanja.** Trenutno sve naše pitanje su tekstualne. Šta je sa znanjem koje se najbolje pokazuje *praktično* — kao kucanje koda, crtanje dijagrama? Sistem može da generiše *zadatke koji nisu MCQ* — open-ended kodiranje, dijagrami — ali procena postaje teža. Aktivno istraživanje.

**Q4. Stabilnost vision modela kod custom UI.** Vec smo videli — vision modeli rade odlično za standardne aplikacije, slabije za custom. Potencijalno rešenje: integracija sa **accessibility API-jima** — Microsoft UI Automation, Windows aria, koji daju strukturisan opis ekrana umesto da se oslanjamo isključivo na vision.

**Q5. Skaliranje na više predmeta.** Da li se ontologija prenosi između domena, ili je svaki predmet svoja zasebna TBox? Trenutno — svaka lekcija ima sopstvenu ontologiju. Skalira li to na ceo program studija? Pitanje koje nismo eksperimentalno odgovorili.

`[prelaz]`

---

## (48) Naredni koraci — Stefan (~3 min)

**Stefan:**
Konkretni naredni koraci — kratki i duži rok.

`[pokaži levo kratki rok, desno duži rok]`

**Kratki rok — sledeća tri do šest meseci.**

Spojiti tri repoziti u jedan monorepo. Tehnički — github workspace ili git submodules, sa zajedničkim CI pipeline-om.

Definisati API kontrakte između slojeva. Konkretno — JSON schema za izlaze Course Planner-a, OpenAPI specifikacija za servise.

Zajednička ontološka osnova — deljena `seed_ontology.ttl` koju svaki sloj proširuje.

Demo end-to-end pipe za jedan predmet. Da prikažemo na konferenciji.

Konferencija — **Sinteza 2026** ili **YuInfo 2026**. Već imamo jedan rad publikovan. Cilj: drugi rad o integraciji.

**Duži rok — godinu i više.**

Pilot u stvarnoj učionici. Kontrola, baseline, statistika. Najveća investicija — ali i najveći dokaz.

Adaptivno biranje pitanja — DRL prototip.

Personalizacija puta učenja — Recommender prototip.

Multimodalna procena — code submissions, dijagrami, audio.

Skaliranje na više predmeta — operativni sistemi, baze, algoritmi, mašinsko učenje.

`[pauza]`

**Etika i privatnost.** Posebna tema koja se *uvek* mora pomenuti kad se priča o AI u obrazovanju.

Lokalni LLM kao osnovna opcija — što sam izbor Uroševog sistema reflektuje. Pravo nastavnika na konačnu reč. Transparentnost — student vidi *zašto* je sistem doneo neku odluku — kroz source_line, anchor, audit trail.

I — što je sada relevantno — **EU AI Act**, koji je u 2025 stupio na snagu, eksplicitno klasifikuje obrazovne AI sisteme kao *high-risk*. To znači — strože transparentnosti, obavezna audit trail, ljudski oversight. Ono što sa našim sistemom imamo prirodno, zbog dizajna sa ontology grounding-om — to bi za druge sisteme bila *posebna investicija*. Naš dizajn je *poklon* za AI Act compliance.

`[prelaz]`

---

## (49) Zaključak — Uroš (~3 min)

**Uroš:**
Zaključak. Vraćamo se na poruku iz uvoda.

`[pokaži 4 kartice]`

**Šta smo pokazali.**

**LLM sam nije dovoljan.** Halucinacije, generičnost, pogrešan kognitivni nivo, odsustvo formalne strukture. Nije pitanje *da li* LLM može da generiše tekst — pitanje je da li može da generiše *pouzdan, pedagoški kalibrisan, ankorisan tekst.* Sam — ne može.

**Pedagoška taksonomija kalibriše.** Bloom drives the WHAT, SOLO drives the HOW DEEP. Ugrađene u prompt, ove taksonomije transformišu LLM iz generatora teksta u pedagoški kalibrisan alat. Lukin projekat to pokazuje za strukturu kursa, moj za pitanja, Stefanov za zadatke u video tutorijalima.

**Ontologija filtrira.** OWL i SPARQL kao formalni ugovor između LLM-a i sistema. Source_line kao anti-halucinacija. Ontology anchor kao kontekst za pitanja. SPARQL validacija kao pre-execution check za planove. Tri komplementarna mehanizma, jedan za svaku vrstu halucinacije.

**Human-in-the-loop ostaje obavezan.** Naročito na visokim SOLO nivoima — gde kvalitet sistemski opada, i gde nastavnikova procena postaje *nezamenljiva*. Sistem snima 80% mehanike, nastavnik odlučuje na 20% gde se najviše vredi.

`[pauza]`

I konačno — *zašto je sve ovo „više od zbira svojih delova"*?

Jer su tri elementa — LLM, ontologija, taksonomija — *međusobno komplementarni*. LLM bez ontologije halucinira. Ontologija bez LLM-a nema sadržaja. Oba bez pedagoške taksonomije proizvode strukturalno tačan ali pedagoški pogrešan output.

Spojeni — postaju sistem koji *zaista* može da podrži nastavu. Ne da je zameni. Da je pojača.

`[prelaz]`

---

## (50) Zaključna misao (citat) — Uroš (~1 min)

**Uroš:**

`[centralni citat na ekranu]`

> *„LLM + ontologija + pedagoška taksonomija = pouzdano, pedagoški kalibrisano ITS okruženje koje je više od zbira svojih delova."*

To je ključna poruka prezentacije. Ako biste zaboravili sve ostalo — *ne zaboravite ovu rečenicu.* Ona je centralna teza, hipoteza, pravac istraživanja i predlog metodologije sve u jednom.

`[prelaz]`

---

## (51) Hvala — Svi (~1 min)

**Svi (Uroš počinje):**

Hvala na pažnji.

**Stefan:** Naročito hvala profesorima Simoni Prokić i Aleksandru Kovačeviću na vođenju kroz ovaj predmet, kao i profesorki Vesin koja je vodila SOTIS i pod čijim mentorstvom su naša tri projekta i nastala.

**Luka:** Otvoreni smo za pitanja i diskusiju.

`[duga pauza]`

---

## (52)–(53) Literatura — Uroš (~2 min, brzo)

**Uroš:**
Za one koji žele da dublje uđu — literatura je organizovana u šest kategorija. Sve reference su na slajdovima, neću da ih čitam jednu po jednu — ali da naglasim ono što je *najviše uticalo* na naš pristup.

**Pedagoške taksonomije** — Biggs i Collis 1982 za SOLO, Anderson i Krathwohl 2001 za revidirani Bloom, Lister 2006 za primenu SOLO u programiranju.

**LLM u obrazovanju** — Kasneci 2023 i Yan 2024 za pregled rizika, Wang 2024 za pregled stanja.

**Automatsko generisanje pitanja** — Kurdi 2020 za sistematski pregled, Liang 2018 za distractor generation.

**Ontologije i Semantic Web** — Vesin 2013 koji je domaći rad iz naše oblasti, Hogan 2021 za knowledge graphs.

**ITS** — Mousavinasab 2021, Lin 2023.

**Multimodalni agenti** — Anthropic Computer Use 2024, OSWorld benchmark Xie 2024.

I dodatne reference koje su od kolega dobile preporuku ili koje smo otkrili u toku rada — Bloom 1984 čuveni „2 Sigma Problem" o jedan-na-jedan tutorstvu kao gornjoj granici onoga što ITS treba da imitira; Kestin 2024 Harvard studija o LLM tutoru koji prevazilazi aktivno-učeničku nastavu; VanLehn 2011 systematic review koji pokazuje da klasični ITS-ovi postižu efektivnost gotovo jednaku ljudskim tutorima.

`[pauza]`

Sve reference su skorije od 2015, osim klasičnih anchor-radova — Bloom 1984, Biggs i Collis 1982, Anderson i Krathwohl 2001 — koji su standardne reference u svakoj diskusiji o pedagoškim taksonomijama.

Hvala još jedanput.

`[kraj]`

---

# Dodaci za prezentera

## Napomene o tempu

- Ukupno: ~10 000–11 000 reči ≈ 75 minuta sa pauzama
- Ako uđete u 90 min, dozvoljeno je da malo proširite Segmente A i C
- Ako ste ispod 60 min, raspoređuje se vreme posle Stefanovog dela u Segmentu B (validacija) — možete dodati live SPARQL primer

## Rotacija govornika

- **Uroš:** otvaranje, tim, predmet, agenda, SOLO, Bloom vs SOLO, dva-pass EA, source_line, ontology anchor, evaluacija SOLO, primeri SOLO, ključne poruke, zaključak, literatura
- **Stefan:** Segment 0 (problem, LLM otvara vrata), dekompozicija zadataka, OWL, RDF/SPARQL, OWL plana, validacija, robustnost UI, naredni koraci, otvorena pitanja, prompt vs finetune (alternativno)
- **Luka:** vizija pipeline-a, tri komponente, struktura kursa, prompt template, JSON serijalizacija, Bloom pokrivenost, integrisana arhitektura, tok podataka, Segment D divider

## Pitanja koja možete očekivati

1. „Zašto Qwen 2.5 a ne Llama?" — *odgovor*: bolji JSON output, bolja srpska podrška, niži VRAM zahtev za 14B
2. „Da li ste merili međuocenjivačku slaganje?" — *odgovor*: u pilotu — Cohen's kappa nije izračunato, ostaje za buduće
3. „Šta sa GPT-4?" — *odgovor*: nismo eksplicitno testirali, ali Mollick & Mollick rad pokazuje da GPT-4 ne menja strukturni problem; pomaže na surface kvalitet ali ne na pedagošku kalibraciju
4. „Kako bi ovo radilo na ne-IT predmetima?" — *odgovor*: pipeline je domen-agnostičan, ontology multi-pass extraction nije specifična za programiranje; testiranje na drugim domenima je deo budućeg rada
5. „Etika korišćenja studentskih podataka za adaptivni RL?" — *odgovor*: lokalni LLM kao baseline, federated learning kao long-term opcija, eksplicitan opt-in
