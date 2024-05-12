import csv
import time

#Tienes que quitar todas las celdas convinadas
#Tienes que quitar las columnas vacias dentro del archivo
#Tienes que convertir todo a num (opcional)
#Tienes que remplazar todas las ñ por n
#Tienes que guardar el archivo como cedulas_comas.csv (separado por comas)
#tienes que meterte al csv y poner todo en numeros

def get_unidades(file_csv):
    unidad_dict = {}
    unidad = ''
    initialRow = 0

    with open(file_csv, 'r', encoding='latin-1') as file:
        csvreader = csv.reader(file)
        rows = list(csvreader)
        for i, row in enumerate(rows):
            for value in row:
                if value.startswith("Unidad") and unidad == '' and initialRow == 0:
                    unidad = value
                    initialRow = i + 3
                    break

                if value.startswith("Unidad") and unidad != '' and initialRow != 0:
                    unidad_dict[unidad] = {'initial': initialRow, 'final': i - 3}
                    unidad = value
                    initialRow = i + 3
                    break

        if unidad != '' and initialRow != 0:
            unidad_dict[unidad] = {'initial': initialRow, 'final': len(rows) - 6}

    return unidad_dict

def get_sales(file_csv, unidad_dict):
    mod_rows = []
    with open(file_csv, 'r', encoding='latin-1') as file:
        csvreader = csv.reader(file)
        rows = list(csvreader)

        for key, values in unidad_dict.items():
            start = values['initial']
            end = values['final']

            for i in range(start, end, 2):
                current_row = rows[i]
                previous = rows[i-1]
                previous[0] = key.replace("Unidad:", "").strip()
                
                if current_row[52] == '0.00':
                    previous[52] = current_row[52] #ARREGLAR ESTO, IDENTIFICA LA ULTIMA SEMANA

                for index, value in enumerate(current_row):         
                    if value != '0.00' and value != '':
                        amount = float(value)
                        if amount < 0:
                            if previous[index] == '':
                                previous[index] += str(amount)
                            else:
                                previous[index] = str(round(float(previous[index]) + amount, 2))
                        else:
                            previous[index] += str(amount)

                    # print(previous)
                    # print(f"   Fila: {i+1}")
                    # print(f"   Indice: {index+1}")
                    # print(f"   Cantidad a cambiar: {amount}")
                    # print("------")
                mod_rows.append(previous)
            # mod_rows.append('')#quitar para reto
    return mod_rows

def new_csv(file_csv, mod_rows):
    new_csv = file_csv.replace('.csv', '_mod.csv')

    with open(new_csv, 'w', newline='', encoding='latin-1') as new_file:
        csvwriter = csv.writer(new_file)
        csvwriter.writerows(mod_rows)

    print(f"Se ha guardado el archivo modificado en: {new_csv}")
                
# Registra el tiempo de inicio
start_time = time.time()

file_csv = '../INPUT/rpt_VentasComisionables_SinCorte.csv'
unidad_dict = get_unidades(file_csv)

print("Unidades encontradas:")
for key, values in unidad_dict.items():
    print(f"{key}: [{values['initial']}, {values['final']}]")

mod_rows = get_sales(file_csv, unidad_dict)
new_csv(file_csv, mod_rows)

end_time = time.time()
elapsed_time = end_time - start_time
print(f"Tiempo de ejecución: {elapsed_time} segundos")