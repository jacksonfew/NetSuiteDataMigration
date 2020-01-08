from General import excel_list, excelcreate, multiexcelwrite


def extractdic(data, mdic):
    head = data[0]
    newhead = []
    for x in head:
        if x in list(mdic.keys()):
            newhead.append(x)
    dic = {'External ID': newhead}
    for row in data[1:]:
        result = []
        for x in newhead:
            i = head.index(x)
            result.append(row[i])
        dic[row[0]] = result

    return dic

def mapdic(fielddata):
    dic = {}
    for row in fielddata[1:]:
        dic[row[1]] = row[0]

    return dic

def appenddata(data, extdic, mapdic):
    head = data[0]
    for x in extdic['External ID']:
        if x in list(mapdic.keys()):
            head.append(mapdic[x])
    full = [head]
    for row in data[1:]:
        nrow = row.copy()
        try:
            dicv = extdic[row[0]]
        except KeyError:
            dicv = []
        for x in dicv:
            nrow.append(x)

        full.append(nrow)

    return full

def textfile(data):
    rows = []
    for row in data:
        r = row.join(',')
        r = r + ' \n'
        rows.append(r)

    return rows

if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite Integration\\Sales Orders 1.6.20\\'

    mapfile = 'Scenarios 1-7 Missing Fields.xlsx'
    writefile =  'Scenarios 1 3 and 8 - 1.7.20 Add Missing Fields.xlsx'
    sheet = 'Scenario 1&3'

    writefile = 'Scenario 7 - Transform 1.4.20 Add Missing Fields.xlsx'
    sheet = 'MASTER'

    data = excel_list(writefile, sheet, folder)
    mdic = mapdic(excel_list(mapfile, 'Map', folder))
    ext = excel_list(mapfile, 'Extract', folder)
    extdic = extractdic(ext, mdic)



    full = appenddata(data, extdic, mdic)
    print(full)

    wdic = {'Sheet1': full, 'Sheet2': [['Jackson Test']]}

    newfile = 'Scenario 7 New Fields.xlsx'



    txt = open('7.txt', 'w')
    textlist = textfile(full)
    txt.writelines(textlist)
    # txt.write('Jackson')
    txt.close()

    excelcreate(full, newfile, 'Sheet1', folder)

    # multiexcelwrite(newfile, wdic, folder)

    # excelcreate([['Jackson Test']], newfile, 'Sheet1', folder)