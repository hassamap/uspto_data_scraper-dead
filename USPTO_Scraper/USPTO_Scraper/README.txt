USPTO Scraper - how to run
==========================

1. Install Python 3 from https://www.python.org/downloads/
   During install, tick "Add python.exe to PATH".
2. Google Chrome or Microsoft Edge is recommended. If neither is installed,
   the first run installs a built-in browser instead.
3. Double-click setup.bat ONCE. It installs everything the scraper needs
   (takes a few minutes). Wait for "Setup complete!".
4. Double-click run.bat to start the scraper every time.
   (If you skip setup.bat, run.bat will still set things up on its first run.)
   Your browser opens to http://127.0.0.1:8765/ automatically.
5. Choose your Excel (.xlsx) file, pick the browser and VPN options, and click Start.
6. When finished, click Download to get the results workbook.

Notes
- Keep the black window open while using the scraper. Closing it stops the app.
- Progress is saved in scraper_state.sqlite3. Re-open the same file later to resume.
- Use "Clear saved data" in the app to delete the saved progress.
- VPN option: CyberGhost must be installed in Chrome and connected to the
  United States. The scraper checks this before every request and stops if it is not.
