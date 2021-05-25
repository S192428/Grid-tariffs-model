All data are in kW, kWh or €/kWh

Done: 
	- H price extracted from GAMS 2016, changed to 2020 prices using the inflation sheet
	- EL price are extracted from GAMS and are the average value for DK1 and DK2 (2016). Corrected with inflation 
	- Inflation sources in the file
	- Heat demand from Philipp, average m2 to adjust the residential profiles (see screenshot...)
	- Intenisities from GAMS 2016, program to compute it in the model folder
	- Taxes (TO STILL CHECK WITH PHILIPP AND CLAIRE), reference document in the main folder (taxing energy in denmark)
	- DH parameters --> arbitrary, and not particularly relevant
	- Grid parameters --> same as for DH
	- Gas price --> not fluctuant
	- Solar CF details are in the excel sheets in the folder SolarCF, plus de sources below (still to finish). Extracted from GAMS and processed with code
	- PV_par (Capital costs = (Capital per MW + OM fixed per MW)*inflation. No O&M costs. Small residential panels from data catalogue)
	- P2H_par (Residential heat pumps in the future Danish energy system (ref in Mendeley) --> air sources HP are the cheapest). Data for the small 1MW air source. No information on year of cost, so considered to be 2020 costs. Same than before, Cap = Inv + FixedOM
	- Boiler: natural gas fired (for DH --> relevance). Considered that the costs are in 2015€ so added the inflation
	- Battery: Li-ion batteries. Depth of discharge and max charge are ~ random. Charging and discharging capacities are estimated regarding the output and input capacities anc compared to the value of the storage capacity, then scaled in %. If higher, capped at 100%
	- Heat storage: Small scale hot water tank (steel). Same charging and discharging capacity than for the battery. Cap = Inv + Fixed OM (scaled to 1h for 1 tank). Inflation
	- Sets kept as it is



TO DO:
	- Electricity profiles




Sources:
	- EL price: https://www.nordpoolgroup.com/historical-market-data/
	- Tech data: Technology catalogue for energy production and storage
	- Gas price: https://www.globalpetrolprices.com/Denmark/natural_gas_prices/ --> check for time series
	- Solar capacity right now is 900 in 2017 (Danish outlook https://ens.dk/sites/ens.dk/files/Analyser/deco19.pdf) + 95.7 + 85 https://www.statista.com/statistics/497380/installed-photovoltaic-capacity-denmark/
	- Capacity in 2016 (bookmark Energy in 2016 in denmark)
	- Gas intensity from Philipp Excel file (in data Adam, folder about the techs, in the Intensity excel sheet)
	- Carbon price --> from conference paper (Word) to justify (17 after pandemic, over 23 otherwise, so 20 as an average). 