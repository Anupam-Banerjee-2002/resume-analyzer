# IILM CareerReady — Resume Analyzer

IILM CareerReady is a student-focused resume analysis platform for IILM University. Students can upload a PDF or DOCX resume to receive career-direction guidance, extracted skills, education and experience insights, and a relevance score. The Flask backend also provides role-based tools for HR, HOD, and ADMIN users.

## Custom model (this build)

The ML classifier was retrained on **460 real, labeled resumes** (`data/resume_dataset.csv`)
across **10 tech job categories**, replacing the original 5-category synthetic-phrase model:

- Java Developer
- Testing
- DevOps Engineer
- Python Developer
- Web Designing
- ETL Developer
- Blockchain
- SAP Developer
- Automation Testing
- DotNet Developer

Two model types were benchmarked on this dataset — **Linear SVC** (calibrated for
probabilities) and **Random Forest** — both scored ~99–100% cross-validated accuracy
(the categories are highly distinct, so this is expected for this dataset, not a sign
of overfitting to noise). The app ships with the **Calibrated LinearSVC** pipeline,
matching the original architecture.

The model is trained automatically from `data/resume_dataset.csv` every time `app.py`
starts (takes a few seconds). To retrain with more/updated data:
- Replace or extend `data/resume_dataset.csv` (columns: `text`, `label`), or
- Use the existing admin endpoints `/admin/dataset/add`, `/admin/dataset/upload`
  (CSV with `text`,`label` columns), then `/admin/dataset/retrain`.

Keyword lists for keyword-scoring (`/keywords`) were also updated to match the new
10 categories with relevant tech keywords per role.

## Run

```bash
python app.py
```

- **Website and API**: `http://127.0.0.1:5000` (the Flask root serves `INDEX.html`)
- **Optional static UI**: `http://127.0.0.1:8000/INDEX.html`
- Uploaded resumes are stored under `uploads/resumes/`
- SQLite DB file: `app.db` (in the project folder by default)

The platform was built for the IILM University student career journey, with product and engineering contributions from **Anupam Banerjee (Frontend · Backend)**.

Institutional context and the logo used in the interface are sourced from the official [IILM University website](https://iilm.edu/). The interface references IILM’s Gurugram, Greater Noida, and Lodhi Road locations and links students back to the university site for official information.

## Default admin (auto-seeded)

If the DB is empty on first run, one ADMIN user is created:

- **Email**: `admin@iilm.local`
- **Password**: `admin@IILM2025`

You can override via env vars:

- `DEFAULT_ADMIN_EMAIL`
- `DEFAULT_ADMIN_PASSWORD`

## Example flow (login → predict → list → download)

```bash
# Login (sets session cookie)
curl -i -c cookies.txt -H "Content-Type: application/json" \
  -d "{\"email\":\"admin@iilm.local\",\"password\":\"admin@IILM2025\"}" \
  http://127.0.0.1:5000/auth/login

# Single predict (PDF/DOCX only) - keeps existing upload field name "resume"
curl -i -b cookies.txt -F "resume=@sample.pdf" http://127.0.0.1:5000/predict

# Bulk predict (PDF/DOCX/ZIP containing PDF/DOCX) - keeps existing field name "resumes"
curl -i -b cookies.txt -F "resumes=@bulk.zip" http://127.0.0.1:5000/predict/bulk

# List resumes (DB-backed)
curl -s -b cookies.txt "http://127.0.0.1:5000/resumes?page=1&page_size=25"

# Download a resume by id
curl -i -b cookies.txt -L "http://127.0.0.1:5000/resumes/<ID>/download" -o downloaded_resume.pdf
```

## Endpoints (grouped by role)

### Auth (all roles)
- `POST /auth/signup`
- `POST /auth/login`
- `POST /auth/logout`
- `GET /auth/me`

### HR
- `POST /predict`
- `POST /predict/bulk`
- `GET /resumes`
- `GET /resumes/<id>/download`
- `DELETE /resumes/<id>`
- **Exports**: `GET /resumes/export/csv|excel|json`
- **Keyword view/update**: `GET /keywords`, `POST /keywords/<department>`

### HOD
- `POST /predict`
- `POST /predict/bulk`
- `GET /resumes` (server forces department scope to the HOD’s department)
- `GET /resumes/<id>/download` (only for own department)
- **Exports**: `GET /resumes/export/csv|excel|json` (scoped to own department)
- **Keyword view**: `GET /keywords` (scoped to own department)

### ADMIN
- Everything HR can do
- **User management**:
  - `GET /users?role=HR|HOD|ADMIN`
  - `POST /users`
  - `DELETE /users/<id>` (cannot delete currently logged-in admin)

### Legacy admin routes (kept working)
- `POST /admin/login` (supports legacy username/password and also ADMIN email/password)
- `POST /admin/logout`
- `GET /admin/status`
- `GET /admin/shortlisted`
- `GET /admin/shortlisted/export/csv|excel|json`
- `GET/POST/DELETE /admin/keywords...` (now HR/ADMIN role-based)
