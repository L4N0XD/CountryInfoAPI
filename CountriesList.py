import json, requests, xlsxwriter # Import modules



response = requests.get("https://restcountries.com/v3.1/all") # Request API Data
print ("Requesting API Data...")
all = json.loads(response.content) # Json Load
print ("Loading JSON...")
workbook = xlsxwriter.Workbook('Downloads/CountriesList.xlsx') # Create a Workbook named "CountriesList" and save in Downloads Path
print ("Creating .xlsx")
worksheet = workbook.add_worksheet("CountriesList") # Create a Worksheet named "CountriesList"


def formatation():
    merge_format = workbook.add_format({'bold': True, 'align': 'center', 'font_color': '#4F4F4F', 'font_size': '16'}) # Formatation for title 
    cell_format = workbook.add_format({'bold': True, 'font_color': '#808080', 'font_size': '12'}) # Formatation for cells
    worksheet.merge_range('A1:D1', 'Countries List', merge_format) # Create title 
    worksheet.write(1, 0, "Name", cell_format) # Create column title Name
    worksheet.write(1, 1, "Capital", cell_format) # Create column title Capital
    worksheet.write(1, 2, "Area", cell_format) # Create column title Area
    worksheet.write(1, 3, "Currencies", cell_format) # Create column title Currencies
    worksheet.set_column(0, 3, 12) # Set column width to 12
    

def names(u):
        try:
            name = (all[u]['name']['common']) # Get country name 
            row = (u + 2)
            column = 0
            worksheet.write(row, column, name) # Write country name to worksheet
            return (name)
        except KeyError:
            name = ("-") # Set "-" for country name
            row = (u + 2)
            column = 0
            worksheet.write(row, column, name) # Write country name to worksheet
            return (name)

def capitals(u):
        try:
            capital = (all[u]['capital'][0]) # Get capital name
            row = (u + 2 )
            column = 1
            worksheet.write(row, column, capital) # Write capital name to worksheet
            return capital        
        except KeyError:
            capital = ("-") # Set "-" for capital name
            row = (u + 2)
            column = 1
            worksheet.write(row, column, capital) # Write capital name to worksheet
            return capital

def areas(u):
        number_format = workbook.add_format({'num_format': '#,##0.00'}) # Defines number format to worksheet
        try:
            area = (all[u]['area']) # Get Area number
            row = (u + 2)
            column = 2
            worksheet.write(row, column, area, number_format) # Write the Area number to worksheet
            return area
        except KeyError:
            area = ("-") # Set "-" for Area number
            row = (u + 2)
            column = 2
            worksheet.write(row, column, area, number_format) # Write the Area number to worksheet
            return area

def currencys(u):
        try:
            getcurrency = (all[u]['currencies']) # Get all currencies
            currencie = list(getcurrency.keys()) # Filter out currencies
            res = str(currencie) # Transform to string
            removed_chars = ['[',"'", ']'] 
            chars = set(removed_chars)
            currencies = ''.join(filter(lambda x: x not in chars, res)) # Remove chars from currencie
            row = (u + 2)
            column = 3
            cell_format = workbook.add_format({'align': 'center'}) 
            worksheet.write(row, column, currencies, cell_format)
            return currencies
        except KeyError: 
            currencies = ("-") # Set "-" for inexistent currencies
            row = (u + 2)
            column = 3
            cell_format = workbook.add_format({'align': 'center'})
            worksheet.write(row, column, currencies, cell_format)
            return currencies

def request():
    for u in range (250):
        print((u), 'Country:', names(u), ' | ', 'Capital: ',capitals(u), ' | ','Area: ', areas(u), ' | ','Currencies: ', currencys(u))
    workbook.close()

formatation()   
request()


