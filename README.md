# Celine & Kiran Wedding Website

This is a Streamlit wedding website built from the files in this folder:

- `app.py` renders the website and RSVP lookup.
- `C&K.xlsx` supplies the wedding copy, event details, registry text, and guest list.
- `assets/` supplies the logo, hero image, and wedding album photos.
- `assets/login-florals-garland.webp` and `assets/login-florals-divider.webp` are the optimized floral artwork used by the login page.
- `assets/album/optimized/` stores the 23 lighter website-ready `.webp` images used by `app.py`.
- `assets/album/originals/` is intentionally kept empty in the deploy copy so GitHub and Streamlit do not carry the full-size source photos.
- `static/CelineandKiran.mp3` supplies the optional background song, which is read directly by `app.py`.
- `.streamlit/config.toml` supplies the Streamlit theme.

The site has a private entry screen. The default password is set in `app.py`, and can be overridden in Streamlit Cloud with a secret named `site_password`.

## Run locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Start the app:

```bash
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this folder to a GitHub repository.
2. Go to `https://share.streamlit.io`.
3. Click `Create app`.
4. Select the GitHub repository and branch.
5. Set the entrypoint file to `app.py`.
6. Click `Deploy`.

Make sure these files are included in GitHub before deploying:

- `app.py`
- `C&K.xlsx`
- `assets/`
- `assets/login-florals-garland.webp`
- `assets/login-florals-divider.webp`
- `assets/album/optimized/`
- `static/CelineandKiran.mp3`
- `.streamlit/config.toml`
