# DDR Report Generator

AI workflow to generate a structured DDR (Detailed Diagnostic Report) from:
- Inspection Report PDF
- Thermal Report PDF

Output is a professional `.docx` report with 7 required DDR sections and relevant images.

## Assignment Fit

This project is designed for the "AI Generalist | Applied AI Builder - DDR Report Generation" task.

It covers:
- extraction of textual observations
- extraction of relevant images from both source PDFs
- merging inspection + thermal context
- handling missing/conflicting details
- generation of a client-ready report structure

## Features

- Reads 2 input PDFs (inspection + thermal)
- Extracts text page-wise using PyMuPDF
- Extracts inspection embedded images and renders thermal pages as images
- Calls Gemini to generate structured DDR JSON
- Builds a formatted DOCX report
- Deduplicates repeated observations/actions
- Handles missing fields with `Not Available` and missing media with `Image Not Available`
- Adds conflict note when contradictory moisture indicators are detected
- Fallback mode: generates DDR structure when Gemini quota/API is unavailable

## Project Structure

- `ddr.py`: full pipeline
- `.env`: API key and runtime config
- `output/`: generated DDR report
- `extracted_images/`: extracted/rendered images
- `ddr_raw.json`: latest DDR JSON payload used for DOCX generation

## Requirements

Install dependencies in your virtual environment:

```bash
pip install pymupdf python-docx google-genai
```

## Configuration (.env)

Create a `.env` file in project root:

```env
GEMINI_API_KEY=your_api_key_here
GEMINI_MODEL=gemini-2.0-flash

# Optional
INSPECTION_PDF=Sample Report.pdf
THERMAL_PDF=Thermal Images.pdf
DDR_OUTPUT_PATH=output/DDR_Report.docx
DDR_ALLOW_FALLBACK=1
```

Notes:
- `GEMINI_MODEL` is optional; code tries fallback model candidates.
- `DDR_ALLOW_FALLBACK=1` allows report generation even if Gemini returns quota/rate-limit errors.

## Run (CLI)

### Default (uses .env / auto-discovery)

```bash
python ddr.py
```

### Explicit inputs

```bash
python ddr.py --inspection "Sample Report.pdf" --thermal "Thermal Images.pdf" --out "output/DDR_Report.docx"
```

### Control fallback behavior

```bash
python ddr.py --allow-fallback
python ddr.py --no-allow-fallback
```

## Run Live API (FastAPI)

Install all dependencies:

```bash
pip install -r requirements.txt
```

Start the API server:

```bash
uvicorn app:app --host 0.0.0.0 --port 8000
```

Open interactive docs:
- http://127.0.0.1:8000/docs

Endpoints:
- `GET /health`: server and config status
- `POST /generate`: upload 2 files as form-data fields:
  - `inspection_pdf` (.pdf)
  - `thermal_pdf` (.pdf)

Response:
- downloadable `DDR_Report.docx`

## Deploy (Render / Railway)

- Build command: `pip install -r requirements.txt`
- Start command: `uvicorn app:app --host 0.0.0.0 --port $PORT`
- Required environment variable: `GEMINI_API_KEY`
- Optional env vars: `GEMINI_MODEL`, `DDR_ALLOW_FALLBACK`

## Output DDR Sections

Generated report contains:
1. Property Issue Summary
2. Area-wise Observations
3. Probable Root Cause
4. Severity Assessment (with reasoning)
5. Recommended Actions
6. Additional Notes
7. Missing or Unclear Information

## Reliability and Limitations

- Gemini quota is project-level. If quota is exhausted, model analysis may fail.
- With fallback enabled, pipeline still generates a valid structured DDR, but deep AI reasoning is reduced.
- Image mapping is heuristic (area text to page match first, index fallback second).


## Future Improvements

- Semantic area-to-image matching using OCR/caption cues
- Better conflict detector with rule scoring
- JSON schema validation and scoring report
- Batch processing for multiple property reports
