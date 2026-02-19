# QuickShip Checker (Lite)

Streamlit web app for estimating PT/AP pallets from a takeoff file.

## Local run
1. Install dependencies:
   - `pip install -r requirements.txt`
2. Start the app:
   - `streamlit run app.py`

## Deployment notes
- Entry point: `app.py`
- Dependencies: `requirements.txt`
- SKU database: `SKU_DATABASE.JSON` (must be deployed alongside the app)
- Theme/server config: `.streamlit/config.toml`

## Configuration
- Update `SKU_DATABASE.JSON` to change PT quantities.
- Optional: add a `logo.png` beside `app.py` to show in the header.
