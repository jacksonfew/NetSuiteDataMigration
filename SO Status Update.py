from General import excel_list, excelcreate

def statusupdate(data, dic):

    head = data[0]
    status_i = head.index('SF_Ascent SO Line Status')
    print(head)
    full = [head]
    for row in data[1:]:
        nrow = row.copy()
        op = row[head.index('SF Order Number')]
        item = row[head.index('Item Sublist: Item - External ID')]
        status = row[status_i]
        if status != 'Partially Packed':
            try:
                status = dic[(op, item)]
            except KeyError:
                status = None
        nrow[status_i] = status
        full.append(nrow)

    return full

def statusdic(ext):
    head = ext[0]

    dic = {}
    for row in ext[1:]:
        op = row[head.index('Sales Order: Sales Order #')]
        item = row[head.index('Product 18 Char ID')]
        status = row[head.index('Status')]

        key = (op, item)
        dic[key] = status

    return dic

if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite Integration\\Sales Orders 1.6.20\\'
    extfile = 'SO Line Status Extract 1.14.20.xlsx'

    sheet = 'Scenario 9'

    ext = excel_list(extfile, sheet, folder)
    sdic = statusdic(ext)
    file = 'SO Line Status Data 1.14.20.xlsx'

    data = excel_list(file, sheet, folder)

    full = statusupdate(data, sdic)

    writefile = 'SO Line Status Update 1.14.20.xlsx'

    excelcreate(full, writefile, sheet, folder)
