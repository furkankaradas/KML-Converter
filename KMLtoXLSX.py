import xlsxwriter
import json

def readkvl(file_name):
    coordinate = 0

    try:
        all_list = []
        with open(file_name, 'rt', encoding="utf-8") as myfile:
            doc = myfile.read()
    except:
        print("File does not find.")
        exit()

    while True:
        coordinate = doc.find('<Placemark>', coordinate + 1)
        if (coordinate == -1):
            break
        else:
            coordinate_start_name = doc.find('<name>', coordinate + 1)
            coordinate_start_name = coordinate_start_name + 6
            coordinate_end_name = doc.find('</name>', coordinate_start_name)
            name = doc[coordinate_start_name: coordinate_end_name]
            coordinate_start_description = doc.find('<description>', coordinate_end_name)
            coordinate_start_description = coordinate_start_description + 13
            coordinate_end_coordinate1 = doc.find('\n', coordinate_start_description)
            description1 = doc[coordinate_start_description: coordinate_end_coordinate1]
            coordinate_end_description = doc.find('</description>', coordinate_start_description)
            description2 = doc[coordinate_end_coordinate1 + 1: coordinate_end_description]
            coordinate_start_coordinate = doc.find('<coordinates>', coordinate_end_description + 1)
            coordinate_start_coordinate = coordinate_start_coordinate + 13
            coordinate_end_coordinate = doc.find('</coordinates>', coordinate_start_coordinate)
            coordinate_maps = doc[coordinate_start_coordinate: coordinate_end_coordinate]
            coordinate_maps = coordinate_maps.lstrip()
            for_space = coordinate_maps.split(' ')
            size = len(for_space)
            new_coordinate = []
            for i in range(size):
                x = for_space[i].split(',')
                for j in range(len(x)):
                    new_coordinate.append(x[j])
            all_list.append(description1)
            all_list.append(name)
            all_list.append(description2)
            all_list.append(new_coordinate[1])
            all_list.append(new_coordinate[0])
            all_list.append(new_coordinate[4])
            all_list.append(new_coordinate[3])
    return all_list

def writexlsx(all_list):
    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True})
    row = 1
    column = 0

    worksheet.write('A1', 'İstasyon Kodu', bold)
    worksheet.write('B1', 'Tanımı', bold)
    worksheet.write('C1', 'Bölgesi', bold)
    worksheet.write('D1', 'Ist. Bas. Kordinatı Enlem', bold)
    worksheet.write('E1', 'Ist. Bas. Kordinatı Boylam', bold)
    worksheet.write('F1', 'Ist. Bitiş. Kordinatı Enlem', bold)
    worksheet.write('G1', 'Ist. Bitiş. Kordinatı Boylam', bold)

    for i in range(len(all_list)):
        worksheet.write(row, column, all_list[i])
        column = column + 1
        if((i+1) % 7 == 0):
            row = row + 1
            column = 0
    workbook.close()

def write_json(all_list):
    size = 0
    temp = {}
    x = 0
    y = 6
    while(True):
        temp[size] = all_list[x:y + 1]
        size = size + 1
        x = y + 1
        y = y + 6 + 1
        list_size = len(all_list) / 7
        if(list_size == size):
            break
    data = {}
    data['coordinates'] = []
    for i in range(size):
        data['coordinates'].append({
            'Istasyon_Kodu': temp[i][0],
            'Tanimi': temp[i][1],
            'Bolgesi': temp[i][2],
            'Baslangic_Enlem': temp[i][3],
            'Baslangic_Boylam': temp[i][4],
            'Bitis_Enlem': temp[i][5],
            'Bitis_Boylam': temp[i][6]
        })
    with open('data.json', 'w', encoding='utf-8') as outfile:
        json.dump(data, outfile, ensure_ascii=False)

def display():
    print("Converter from KML to EXCEL or JSON")
    while (True):
        print(15 * "*")
        ch = input("1.Converter\n2.Exit\nPlease Enter Choose:")
        if (ch == "1"):
            file_name = input("Please enter document name (Example: example.kml) :")
            all_list = readkvl(file_name)
            choose = input("Please enter converter file type (json or excel):")
            if (choose == "json"):
                write_json(all_list)
                return
            elif (choose == "excel"):
                writexlsx(all_list)
                return
            else:
                print("Wrong Type. (Please write json or excel)")
        elif (ch == "2"):
            return
        else:
            print("Wrong Type.")

def main():
    display()

if __name__ == "__main__":
    main()