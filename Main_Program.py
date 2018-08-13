import openpyxl
import Write_File

file1 = r'C:\Python_Scripts\Get_Site_Info_By_OID\Master_Sheet.xlsx'
file2 = r'C:\Python_Scripts\Get_Site_Info_By_OID\SiteLookup.xlsx'
wb1 = openpyxl.load_workbook(file1)
wb2 = openpyxl.load_workbook(file2)
sht1 = wb1.get_sheet_by_name('Facilities')
sht2 = wb2.get_sheet_by_name('Sheet1')
mylist = []
adv_data = []
fin_data = []

def get_ADV_info(mylist):
    for i in mylist:
        x = 3
        tcell = sht1['A3'].value
        while tcell != None:
            cellA = str(sht1['E' + str(x)].value)
            if cellA == i:
                data = [i, str(sht1['A' + str(x)].value), str(sht1['Z' + str(x)].value)]
                adv_data.append(data)
                data = []
                break
            x += 1
            tcell = sht1['A' + str(x)].value

def get_myList():
    x = 2
    tcell = sht2['A1'].value
    while tcell != None:
        cellA = str(sht2['A' + str(x)].value)
        mylist.append(cellA)
        x += 1
        tcell = sht2['A' + str(x)].value

def update_list(mylist, adv_data):
    X = 0
    y = 0
    i = 0
    for i in mylist:
        fac = adv_data[y]
        if i == fac[X]:
            fin_data.append(fac)
            y += 1
        else:
            data = [i, 'None', 'None']
            fin_data.append(data)


def main_run():
    get_myList()
    get_ADV_info(mylist)
    wb1.save(file1)
    wb2.save(file2)
    update_list(mylist, adv_data)
    select = Write_File.data_write_to_file()
    Write_File.write_to_sheet(select, fin_data)

main_run()
