from General import excel_list, excelcreate, multiexcelwrite


def Dupes(data):

    head = data[0]
    itemi = head.index('Item: Item Number')
    desi = head.index('Product: Product Name')
    mpni = head.index('Product: MPN')

    items = {'Item Number': ['Description 1', 'Description 2']}
    mpns = {'MPN': ['Item 1', 'Item 2']}
    nompn = [['Item Number', 'Description']]

    for row in data[1:]:
        item = row[itemi]
        des = row[desi]
        mpn = row[mpni]

        if item not in list(items.keys()):
            items[item] = [des]
        else:
            items[item].append(des)
        if mpn != None:
            if mpn not in list(mpns.keys()):
                mpns[mpn] = [item]
            else:
                mpns[mpn].append(item)
        else:
            nompn.append([item, des])

    idupe = {}
    mdupe = {}

    for x in list(items.keys()):
        if len(items[x]) > 1:
            idupe[x] = items[x]

    for x in list(mpns.keys()):
        if len(mpns[x]) > 1:
            mdupe[x] = mpns[x]

    return idupe, mdupe, nompn


def dictolist(dic):

    full = []
    for x in list(dic.keys()):
        row = dic[x]
        row.insert(0, x)
        full.append(row)

    return full


if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite\\'
    file = 'MPN Dupe Check.xlsx'

    report = excel_list(file, 'Report', folder)

    dupes = Dupes(report)

    itemdic = dupes[0]
    mpndic = dupes[1]
    nompn = dupes[2]

    mpn = dictolist(mpndic)

    sheetdic = {'MPN': mpn, 'No MPN': nompn}

    multiexcelwrite(file, sheetdic, folder)

    # excelcreate(mpn, file, 'MPN', folder)