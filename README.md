[🇷🇺 Русский](README.ru.md)
---

# Energy Drink Collection Manager
A cross-platform Flet application for tracking, cataloging, and analyzing an energy drink collection.

## Key Features
- **Collection management:** list view, quick search, sorting by brand, price, and rating, detailed item card, swipe to delete an entry.
- **Beverage details:** recording the ingredients (caffeine, taurine, sugar), volume, price, and place of purchase; personal rating; selecting a custom card color; and attaching photos.
- **Flexible field configuration:** enabling/disabling standard parameters and creating custom fields.
- **Statistics:** calculation of total volume, funds spent, average rating, and the top 5 popular brands.
- **Imports and exports:**
  - Import from Excel files with column selection and cell color retrieval.
  - Export and import of the database in JSON format.

## Technology Stack
- Python 3.14.7
- Flet 0.86.5 (UI)
- openpyxl 3.1.5 (Excel parsing)

## Installation and Launch
```bash
git clone [https://github.com/your-username/energy-drink-manager.git](https://github.com/your-username/energy-drink-manager.git)
cd energy-drink-manager
pip install flet openpyxl
python main.py
