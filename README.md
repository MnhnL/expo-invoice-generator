# expo-invoice-generator
Generates pdf invoices by commune from csv exports

## Building a .exe on windows
1. Download a .zip and extract somewhere
2. Create a venv: `python -m venv venv`
3. Activate venv: `venv\Scripts\activate.bat`
4. Maybe upgrade pip: `python -m pip install pip --upgrade`
5. Install requirements: `python -m pip install -r requirements.txt`
6. Build: `python setup.py py2exe`
7. Find the .exe in the `dist` directory.
