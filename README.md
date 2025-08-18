# 🐛 RUPS Vergelijkingstool

Deze repository bevat een Streamlit-applicatie voor het vergelijken van RUPS-planningen tussen twee verschillende jaren. De tool is bedoeld om verschillen in maatregelen te identificeren en te visualiseren op basis van geüploade Excel-bestanden.

## ⚙️ Functionaliteit
- Upload twee Excel-bestanden (oud en nieuw) met RUPS-planningen.
- Vergelijk de planningen en identificeer toegevoegde, verwijderde en gewijzigde maatregelen.
- Download het resultaat als een Excel-bestand met drie tabbladen: Vergelijking, Verwijderd, Toegevoegd.

## 🚀 Gebruik
1. Start de applicatie:
	```powershell
	streamlit run compare_rups.py
	```
2. Upload het oude en nieuwe RUPS-bestand via de zijbalk.
3. Selecteer de juiste kolommen voor maatregelnaam en maatregelnr.
4. Klik op "Voer vergelijking uit" om de analyse uit te voeren.
5. Download het resultaat als Excel-bestand.

## 📦 Vereisten
- Python 3.8+
- streamlit
- pandas
- openpyxl
- regex

Installeer de benodigde packages:
```powershell
pip install -r requirements.txt
```

## 📁 Bestanden
- `compare_rups.py`: Hoofdscript voor de Streamlit-app.
- `README.md`: Deze documentatie.
- `requirements.txt`: Lijst van vereiste Python-pakketten.

## 📬 Contact
Voor vragen of suggesties, neem contact op met de beheerder van deze repository.
