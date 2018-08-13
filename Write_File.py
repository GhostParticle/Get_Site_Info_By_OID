import openpyxl


def data_write_to_file():
    print("How would you like this formatted, please chose a number from the following")
    print("1: Shortform (Acronyms Only)")
    print("2: Longform (Acronym and Site)")
    print("3: Full Form (Oracle ID: Acronym: Site Name)")
    select = int(input())
    return select

def write_to_sheet(select, fin_data):
    file = r'C:\Python_Scripts\Get_Site_Info_By_OID\SiteLookup.xlsx'
    wb = openpyxl.load_workbook(file)
    sht = wb.get_sheet_by_name('Sheet1')
    x = 2
    if select == 1:
        sht['C1'] = 'Advantx Name'
        for i in fin_data:
            print('{}'.format(i[2]))
            if i[2] == 'None':
                sht['C' + str(x)] = 'NA'
                print('NA')
            else:
                sht['C' + str(x)] = i[2]
            x += 1
        wb.save(file)
    if select == 2:
        sht['C1'] = 'Advantx Name'
        sht['D1'] = 'Site Name'
        for i in fin_data:
            print('{} : {}'.format(i[2], i[1]))
            if i[2] == None:
                sht['C' + str(x)] = 'NA'
                sht['D' + str(x)] = 'NA'
            else:
                sht['C' + str(x)] = i[2]
                sht['D' + str(x)] = i[1]
    if select == 3:
        sht['C1'] = 'Advantx Name'
        sht['D1'] = 'Site Name'
        sht['E1'] = "Oracle ID"
        for i in fin_data:
            print('{} : {} : {}'.format(i[2], i[1], i[0]))
            if i[2] == 'None':
                sht['C' + str(x)] = 'NA'
                sht['D' + str(x)] = 'NA'
                sht['E' + str(x)] = 'NA'
            else:
                sht['C' + str(x)] = i[2]
                sht['D' + str(x)] = i[1]
                sht['E' + str(x)] = i[0]
            print(len(i))
            print(i[2])
            print(i[1])
            print(i[0])
            x += 1
        wb.save(file)

