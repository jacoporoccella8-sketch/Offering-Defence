# NTT DATA – Defence Market Scraper

Script Python che monitora quotidianamente i player del mercato Difesa italiano,
genera un **Excel strutturato** e invia una **mail di riepilogo** con il delta
rispetto all'esecuzione precedente.

---

## Struttura del progetto

```
defence-market-scraper/
├── defence_market_scraper.py   ← script principale
├── requirements.txt
├── .env.example                ← copia in .env e compila
├── .github/
│   └── workflows/
│       └── daily_scraper.yml   ← automazione GitHub Actions
├── data/
│   └── history.json            ← storico automatico (creato dallo script)
└── output/
    └── defence_market_YYYYMMDD.xlsx
```

---

## 1. Installazione in locale

```bash
# 1. Clona il repo (o scarica i file)
git clone https://github.com/TUO_USERNAME/defence-market-scraper.git
cd defence-market-scraper

# 2. Installa le dipendenze
pip install -r requirements.txt

# 3. Configura le variabili d'ambiente
cp .env.example .env
# → apri .env con un editor di testo e compila SMTP_USER, SMTP_PASS, ecc.

# 4. Prima esecuzione
python defence_market_scraper.py --once
```

L'Excel viene salvato nella cartella `output/` e la mail viene inviata
a `jacopo.roccella@nttdata.com`.

---

## 2. Esecuzione schedulata in locale

```bash
python defence_market_scraper.py
# Lo script gira in loop e si attiva ogni giorno alle 07:30
```

---

## 3. Output Excel

Il file generato contiene **3 fogli**:

| Foglio | Contenuto |
|---|---|
| **Dashboard** | Tabella principale con tutti i player, keyword rilevate, delta vs ieri |
| **Gare & Procurement** | Dettaglio fonti gare per ogni player |
| **Fonti & Metodologia** | URL sorgente con link cliccabili e tipo di pista |

La colonna **"Delta vs ieri"** segnala (con sfondo giallo) ogni variazione
rispetto all'ultima esecuzione: nuovi importi, keyword apparse/scomparse,
variazioni nella raggiungibilità delle fonti.

---

## 4. Configurare GitHub Actions (guida per neofiti)

> **Cos'è GitHub Actions?** È un sistema di automazione gratuito integrato in
> GitHub: esegue il tuo script su un server remoto secondo una pianificazione,
> senza che tu debba tenere acceso il computer.

### Passo 1 – Crea un account GitHub

1. Vai su [https://github.com](https://github.com) → click **Sign up**
2. Scegli un username, inserisci email e password
3. Verifica la mail

### Passo 2 – Crea un nuovo repository

1. Dopo il login, click sul **"+"** in alto a destra → **New repository**
2. Dai un nome, es. `defence-market-scraper`
3. Seleziona **Private** (per mantenere riservato il codice)
4. Click **Create repository**

### Passo 3 – Carica i file

Opzione A – da browser (più semplice):
1. Nella pagina del repository → **Add file → Upload files**
2. Trascina tutti i file (`.py`, `requirements.txt`, `.env.example`,
   la cartella `.github/`)
3. Click **Commit changes**

Opzione B – da terminale:
```bash
cd defence-market-scraper
git init
git remote add origin https://github.com/TUO_USERNAME/defence-market-scraper.git
git add .
git commit -m "primo commit"
git push -u origin main
```

### Passo 4 – Aggiungi i Secrets (credenziali SMTP)

Non mettere mai password in chiaro nel codice. GitHub usa i **Secrets**:

1. Nel tuo repository → **Settings → Secrets and variables → Actions**
2. Click **New repository secret** per ognuno:

| Nome Secret      | Valore da inserire                    |
|------------------|---------------------------------------|
| `SMTP_HOST`      | `smtp.office365.com`                  |
| `SMTP_PORT`      | `587`                                 |
| `SMTP_USER`      | la tua email NTT DATA                 |
| `SMTP_PASS`      | la tua password (o App Password)      |
| `RECIPIENT_EMAIL`| `jacopo.roccella@nttdata.com`         |

### Passo 5 – Verifica che Actions funzioni

1. Vai su **Actions** (tab in alto nel repository)
2. Dovresti vedere il workflow `Defence Market – Daily Scraper`
3. Click **Run workflow** → **Run workflow** per testarlo subito
4. Se tutto è verde ✅ riceverai la mail e troverai l'Excel negli **Artifacts**

### Passo 6 – Schedule automatico

Il file `.github/workflows/daily_scraper.yml` contiene già:
```yaml
- cron: "30 6 * * *"   # 07:30 CET ogni giorno
```
Non devi fare altro: GitHub eseguirà lo script ogni mattina.

---

## 5. Nota su Office 365 e App Password

Se il tuo account NTT DATA ha **MFA attivo** (molto probabile), la password
normale non funzionerà per SMTP. Devi generare una **App Password**:

- Microsoft 365 → Account personale → Sicurezza → App password
- Oppure chiedi al tuo IT NTT DATA di abilitare SMTP AUTH sul tuo account

---

## 6. Personalizzazioni rapide

| Cosa vuoi cambiare | Dove |
|---|---|
| Orario invio mail | `daily_scraper.yml` → riga `cron:` |
| Aggiungere un player | `PLAYERS` list in `defence_market_scraper.py` |
| Aggiungere una fonte | `"fonti"` del player corrispondente |
| Cambiare destinatario | `.env` → `RECIPIENT_EMAIL` |
