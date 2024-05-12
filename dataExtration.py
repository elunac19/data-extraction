import pandas as pd
import csv
import time

xlsFile = 'rpt_VentasComisionables_SinCorte.xls'
csvFile = 'rpt_VentasComisionables_SinCorte.csv'
csvFinalFile = 'TeamSales.csv'


df = pd.read_excel(xlsFile)
df.to_csv(csvFile, index=False)

print(f"The file '{csvFile}' has been successfully created.")

# Gets the range of cells for each team. 
# Example: Unit: 1/CENTRAL UNIT: [11, 124]. 
# With this we can then iterate over these ranges to extract the values and modify them.
def getTeam():
    team = {}
    teamName = ''
    initialRow = 0

    with open(csvFile, 'r', encoding='latin-1') as file:
        csvreader = csv.reader(file)
        rows = list(csvreader)
        for i, row in enumerate(rows):
            for value in row[:3]: #As the team name is always in the 3rd position, it is convenient to cut the rows to 3 values.  
                if value.startswith("Unidad") and teamName == '' and initialRow == 0:
                    teamName = value
                    initialRow = i + 3
                    break

                if value.startswith("Unidad") and teamName != '' and initialRow != 0:
                    team[teamName] = {'initial': initialRow, 'final': i - 3}
                    teamName = value
                    initialRow = i + 3
                    break

        if teamName != '' and initialRow != 0:
            team[teamName] = {'initial': initialRow, 'final': len(rows) - 6}

    return team

#With the cells range of each team, now the sales are obtain. 
#Since the document has positive and negative sales, due to cancelations.
#Its important to recalculate the sales by merging the 2 rows into one.  
def getTeamSales(team):
    mod_rows = []

    with open(csvFile, 'r', encoding='latin-1') as file:
        csvreader = csv.reader(file)
        maxColumn = len(next(csvreader)) - 1
        rows = list(csvreader)

        for key, values in team.items():
            start = values['initial']
            end = values['final']

            for i in range(start, end, 2):
                salesRow = rows[i]
                cancelsRow = rows[i-1]
                salesRow[0] = key.replace("Unidad: ", "").strip()
                print(salesRow)
                
                # if salesRow[maxColumn] == '0.00':
                #     cancelsRow[maxColumn] = salesRow[maxColumn]

                for index, value in enumerate(salesRow):         
                    if value != '0.00' and value != '0' and value != '':
                        amount = float(value.replace(',', ''))
                        if amount < 0:
                            if cancelsRow[index] == '':
                                cancelsRow[index] += str(amount)
                            else:
                                cancelsRow[index] = str(round(float(cancelsRow[index].replace(',', '')) + amount, 2))

                        else:
                            cancelsRow[index] += str(amount)

                mod_rows.append(cancelsRow)
    return mod_rows


def new_csv(teamSales):
    with open(csvFinalFile, 'w', newline='', encoding='latin-1') as new_file:
        csvwriter = csv.writer(new_file)
        
        csvwriter.writerows(teamSales)

start_time = time.time()

print(f"Getting all teams...")
team = getTeam()

print("Unidades encontradas:")
for key, values in team.items():
    print(f"{key}: [{values['initial']}, {values['final']}]")

# print(f"Getting all team sales...")
# teamSales = getTeamSales(team)
# print(f"Creating new file...")
# new_csv(teamSales)

end_time = time.time()
elapsed_time = end_time - start_time
print(f"Tiempo de ejecución: {elapsed_time} segundos")       