# ProductivityAI Setup

A Windows voice assistant leveraging speech recognition, TTS, local LLM via Ollama, and various automation utilities.

## 1. Prerequisites
- Windows 10/11 (with microphone access enabled)
- Python 3.10+ (recommend 64-bit)
- Administrator rights for some system actions (lock, shutdown)
- Internet connection (for Wikipedia, weather, news, WolframAlpha, Ollama model pull, etc.)

## 2. External Applications / Services
Install or ensure availability of:
- Google Chrome (path used: `C:\Program Files\Google\Chrome\Application\chrome.exe`)
- Discord, Zoom, Notion, MS Office (adjust paths if different on your system)
- Tesseract OCR (for pytesseract) — install from: https://github.com/tesseract-ocr/tesseract. After install, add its install directory to PATH or set `pytesseract.pytesseract.tesseract_cmd`.
- Ollama (for local Llama model): https://ollama.com/  After install run: `ollama pull llama3.2`

Create required API keys:
- OpenWeatherMap API key (replace placeholder in code)
- WolframAlpha App ID (replace placeholder in code)
- News API key (replace placeholder in code if you intend to use news feature)

## 3. Python Environment Setup
```powershell
# From project root
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
```

If `pywin32` post-install script is needed:
```powershell
python Scripts/pywin32_postinstall.py -install
```

## 4. Configure Paths & Keys
Open `ProductivityAI_main.py` and adjust:
- Application shortcut paths (`codePath` variables) to match your system.
- Replace placeholders:
  - OpenWeatherMap: `api_key = "YOUR_KEY"`
  - WolframAlpha: `app_id = "YOUR_APP_ID"`
  - News API: `apiKey": "YOUR_NEWS_API_KEY"`

## 5. (Optional) Tesseract Path
If not on PATH:
```python
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```
Add near the imports.

## 6. Ollama Model
Ensure the model name in `generate_response` exists locally. Example:
```powershell
ollama pull llama3.2
```
List models: `ollama list`
Change model variable if needed.

## 7. Run the Assistant
```powershell
.\.venv\Scripts\Activate.ps1
python ProductivityAI_main.py
```
Grant microphone permission on first run.

## 8. Common Issues
- Microphone not found: verify default recording device & run PowerShell as normal user (not elevated) if needed.
- TTS voice index error: print available voices:
  ```python
  import pyttsx3; e=pyttsx3.init(); print(e.getProperty('voices'))
  ```
  Adjust `voices[1]` if out-of-range.
- Ollama connection error: ensure `ollama serve` service is running (it auto-starts after install) or run manually.
- pytesseract errors: verify Tesseract install path.

## 9. Security Note
This script can execute shutdown/restart and open applications. Review commands before running in a production machine.

## 10. Future Improvements
- Add graceful exit keyword outside "Hello" gating.
- Refactor into modular structure with classes.
- Add hotword detection instead of continuous loop.

---
Adjust as needed for your environment.
