# World Cup Points App - Streamlit Deployment

## Files

- `app.py` - Streamlit app entry point.
- `requirements.txt` - Python dependencies for Streamlit Cloud or other hosts.
- `.streamlit/secrets.toml.example` - Example secrets file. Do not commit real tokens.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## API token setup

The app checks for a token in this order:

1. User input in the Streamlit sidebar.
2. `st.secrets["FOOTBALL_DATA_API_TOKEN"]`.
3. Environment variable `FOOTBALL_DATA_API_TOKEN`.

For Streamlit Community Cloud, add this in **App settings -> Secrets**:

```toml
FOOTBALL_DATA_API_TOKEN = "your_real_token"
```

## Deploy on Streamlit Community Cloud

1. Push `app.py` and `requirements.txt` to a GitHub repository.
2. Go to Streamlit Community Cloud and create a new app from that repo.
3. Set the main file path to `app.py`.
4. Add `FOOTBALL_DATA_API_TOKEN` in app secrets.
5. Deploy.
