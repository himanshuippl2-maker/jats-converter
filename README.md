# DOCX → JATS XML Converter
**PMC-Compliant Medical Journal Article Converter**

## 🚀 Deploy in 5 minutes

### Option 1: Railway (Recommended — Free tier available)
1. [railway.app](https://railway.app) pe account banao
2. "New Project" → "Deploy from GitHub repo" click karo
3. Apna GitHub repo connect karo (in files upload karo)
4. Automatic deploy ho jaayega ✅

### Option 2: Render (Free tier)
1. [render.com](https://render.com) pe account banao
2. "New Web Service" → GitHub repo connect karo
3. `render.yaml` auto-detect ho jaayega
4. Deploy click karo ✅

### Option 3: Heroku
```bash
heroku create your-app-name
git push heroku main
```

---

## 🖥️ Local Run
```bash
pip install -r requirements.txt
python app.py
# Open: http://localhost:5000
```

---

## 🔌 API Usage

### Convert via curl
```bash
curl -X POST https://your-app.railway.app/api/convert \
  -F "file=@article.docx" \
  -F "journal=IP Indian Journal of Clinical and Experimental Dermatology" \
  -F "issn_print=2581-4710" \
  -F "issn_elec=2581-4729" \
  -F "publisher=IP Innovative Publication" \
  -F "doi=10.18231/j.ijced.2025.001" \
  -F "volume=11" -F "issue=4" \
  -F "year=2025" -F "month=12" -F "day=30" \
  -F "fpage=473" -F "lpage=478" \
  -F "article_type=research-article" \
  -F "license=cc-by-nc-4.0" \
  -o output.xml
```

### API Parameters

| Field | Required | Example |
|---|---|---|
| `file` | ✅ | article.docx |
| `journal` | ✅ | IP Indian Journal of... |
| `publisher` | ✅ | IP Innovative Publication |
| `issn_print` | ✅ | 2581-4710 |
| `issn_elec` | ✅ | 2581-4729 |
| `doi` | ✅ | 10.18231/j.xxx.2025.001 |
| `volume` | ✅ | 11 |
| `issue` | ✅ | 4 |
| `year` | ✅ | 2025 |
| `month` | ✓ | 12 |
| `day` | ✓ | 30 |
| `fpage` | ✓ | 473 |
| `lpage` | ✓ | 478 |
| `article_type` | ✓ | research-article |
| `license` | ✓ | cc-by-nc-4.0 |
| `crossref` | ✓ | true/false (default: false) |

### Article Types
- `research-article` — Original Research Article
- `review-article` — Review Article
- `case-report` — Case Report
- `letter` — Letter to Editor
- `editorial` — Editorial
- `brief-report` — Brief Report
- `systematic-review` — Systematic Review

### License Options
- `cc-by-nc-4.0` — CC BY-NC 4.0 (default)
- `cc-by-4.0` — CC BY 4.0
- `cc-by-nc-nd-4.0` — CC BY-NC-ND 4.0

---

## ✅ PMC Compliance
- JATS DTD v1.2
- `pub-date @date-type + @publication-format`
- Structured abstract with `<sec><title><p>`
- `sec-type` on standard sections
- `<permissions>` with CC `<license xlink:href>`
- `<author-notes>` with `<corresp>` + `<fn fn-type>`
- `<floats-group>` after `<back>`
- No empty elements

## 📋 Word File Requirements
Your .docx must use these paragraph styles:
- `Title` — Article title
- `Author Name` — Authors with superscript affiliation numbers
- `Authors affiliation` / `Last Authors affiliation` — Affiliations
- `abstract heading` + `Abstract` — Abstract
- `Keywords` — Keywords and dates
- `Heading 1` / `Heading 2` — Section headings
- `Paragraph 1` / `2nd Para` — Body text
- `Table caption` — Table captions
- `Reference` — References
