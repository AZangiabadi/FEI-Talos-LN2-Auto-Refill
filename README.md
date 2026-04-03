# LN2 Reader Filler

This program is a desktop monitor for the FEI Talos TEM LN2 system. It reads the current liquid nitrogen level from the microscope, watches for low-level conditions, and can automatically or manually trigger a refill through the WebSwitch relay. It also includes a cryo-cycle mode that suppresses refills when monitoring only is desired.

This project can be run with `uv` as a small desktop app.

Before running the app, set the `LN2_WEBSWITCH_BASE_URL` environment variable on your machine to the local relay that controls the solenoid LN2 valve, for example `http://<your-webswitch-ip>`. This value is intentionally not stored in the repository.

## Run with uv

If `uv` is installed:

```powershell
uv sync
uv run ln2-reader-filler
```

You can also run the file directly through `uv`:

```powershell
uv run python main.py
```
<img width="440" height="313" alt="image" src="https://github.com/user-attachments/assets/308e8369-2127-499d-9385-737ad597a105" />


## Install uv

If `uv` is not available on your machine yet, install it first:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then restart your terminal and run the commands above from this folder:

`C:\Users\amira\OneDrive\Documents\LN2_reader_filler`
