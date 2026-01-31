This document outlines the correct procedure for populating the Excel Sheets files connected to the Asset Management Dashboard python code and the procedure to launch teh dashboard.
⚠️ GOLDEN RULE: Never modify the column names (the first header row) in the tabs. The application relies on these exact names to read the data. Changing them will cause the app to crash.

1. Tab: Comptes (Current Portfolio)
This tab represents the real-time "snapshot" of the assets you currently own. It is used to calculate the total portfolio value and current allocation.

	📝 Entry Procedure
Each row corresponds to a specific account observed at a certain date. 
📊 Columns to Fill
Column 				Description 					Format 
Date  				Date of today					DD-MM-YYYY 
Compte 				Name or type of account				(free format)
Montant				Total amount available on the account		 	

2. Tab: Stratégie et Feuille de Route_Bourse (Roadmap of orders)
This tab is the "brain" of the application. It allows you to define your objectives, valuation estimates, and spot undervalued stocks.

	📝 Entry Procedure
This is where you put your buy and sell orders.
Each row corresponds to a specific asset held.
📊 Columns to Fill
Column 				Description 					Format 
Date exécution Ordres 		Date of the last order.				DD-MM-YYYY 
Type d'ordre			Type of the order				"Achat", "Vente"
Compte 				Name or type of account				"PEA" "CTO" (free format)
Catégorie			Categorie of asset				"Action"
Ticker				Unique Stock Symbol (Yahoo Finance)		AAPL, MC.PA, AIR.PA  
Type Ordre			Type of order					"Au marché", "à cours limité"
Nombre d'ordre			Number of stocks bought				
Valeur total Ordres + frais	Total price paid to the broker for the order 	


🟢 Do's
Decimal Separators: Be consistent with your  settings ( comma ,). The code handles conversion, but consistency prevents errors.Yahoo Tickers: For non-US stocks, do not forget the suffix. Use .PA for Paris, .L for London, etc. (e.g., AIR.PA for Airbus).
Clean Data: Remove any empty rows in the middle of your table.
🔴 Don'ts
No Currency Symbols: Do not include currency symbols (€, $) inside the quantite or prix cells. 
Keep them as pure numbers.
Bad: 150 €Good: 150
No Merged Cells: Never merge cells within the data table.🔄 Updating the AppAfter modifying this Sheets file:Close the Google Sheets tab (or ensure auto-save is complete).Open the Dashboard.Use the "Refresh" button (or pull-to-refresh) to load the new data.

3. Code execution 

Once every sheet is fulfilled, launch the file "lanceur.py" and exectute the code, it will open the Dashboard.
