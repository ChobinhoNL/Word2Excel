from wordextractor import wordext
import xlsxwriter

def main():
    waar = str(input("Vul de path van de folder in: "))
    factuurlist = wordext(waar) # Getting a list with the class-data per invoice.
    a = 0
    length = len(factuurlist)

    for i in factuurlist:
        print(i)

    # Create a workbook and add a worksheet.
    wb = xlsxwriter.Workbook(f'{waar}\Admin.xlsx')
    ws = wb.add_worksheet()

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    while length != 0:
        ws.write(row, col, factuurlist[a].datum)
        ws.write(row, col+1, factuurlist[a].fnummer)
        ws.write(row, col+2, factuurlist[a].bedrijf)
        ws.write(row, col+5, factuurlist[a].incl)
        length -= 1
        a += 1
        row += 1


    wb.close()

if __name__ == '__main__':
    main()