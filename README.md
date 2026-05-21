# Celine & Kiran Wedding Website

Streamlit wedding website for Celine and Kiran.

## Project Files

- `app.py` renders the website, password screen, music player, and RSVP lookup.
- `C&K.xlsx` supplies the wedding copy, event details, registry text, guest list, and RSVP response sheet.
- `assets/` supplies the logo and hero image.
- `static/CelineandKiran.mp3` supplies the optional background song.
- `.streamlit/config.toml` supplies the Streamlit theme and enables static file serving.

## Streamlit Deploy

Use `app.py` as the entrypoint file.

Make sure these files are pushed to GitHub:

- `app.py`
- `C&K.xlsx`
- `assets/`
- `static/CelineandKiran.mp3`
- `.streamlit/config.toml`
- `requirements.txt`

The default site password is set in `app.py`. For better privacy, set a Streamlit secret named `site_password`.
