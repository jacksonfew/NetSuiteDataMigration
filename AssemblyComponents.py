from General import excel_list, excelcreate, binaryconvert as bc



def AssCompForm(data, iddic, error=None, iid=True):
    head = data[0]

    # dedic ={}
    compdic = {}
    full = []

    for row in data[1:]:
        item = row[head.index('Item')]
        if error == 'Item':
            error = item

        qty = row[head.index('Quantity')]
        # de = row[head.index('Default?')]
        bomid = row[head.index('Bill of Materials: ID')]
        bomitem = row[head.index('BOM Item')]
        remove = row[head.index('Remove')]
        bom = (bomitem, bomid)



        if remove != True:
            if iid == True:
                try:
                    item = iddic[str(item)]
                except KeyError:
                    item = error
            else:
                item = item
            if bom not in list(compdic.keys()):
                # dedic[bom] = bc(de)
                compdic[bom] = [item, qty]
                full.append([bomitem, bomid])
            else:
                compdic[bom].append(item)
                compdic[bom].append(qty)

    maxlen = 0
    for row in full:
        key = (row[0], row[1])
        # row.append(dedic[key])


        for x in compdic[key]:
            row.append(x)

        L = len(compdic[key])
        if L > maxlen:
            maxlen = L


    newhead = ['BOM Item', 'BOM ID'] #, 'Default']
    for i in range(1, int(maxlen/2)+1):
        txt = 'Assembly Member Sublist '+str(i)
        itemtxt = txt + ': Item'
        qtytxt = txt + ': Item Quantity'
        newhead.append(itemtxt)
        newhead.append(qtytxt)

    full.insert(0, newhead)

    print(len(full))
    return full

def transform(data, custombom):
    head = data[0]
    i = head.index('BOM Item')
    bom = head.index('Bill of Materials: ID')
    de = head.index('Default?')
    re = head.index('Remove')
    date = head.index('Bill of Materials: Created Date')

    itemdic = {}
    bomdic = {'External ID': 'BOM Name'}
    idic = {}

    for row in data[1:]:
        if row[re] != True:
            n = [row[date], row[bom], row[i], bc(row[de])]
            if row[i] not in list(itemdic.keys()):
                itemdic[row[i]] = [n]
            elif n not in itemdic[row[i]]:
                itemdic[row[i]].append(n)

            itemdic[row[i]].sort()

    for item in list(itemdic.keys()):
        n = itemdic[item]
        y = n.copy()
        for row in n:
            i = n.index(row)
            v = i+1
            id = row[1]
            if id not in list(custombom.keys()):
                name = 'BOM_' + str(item) + '_v' + str(v)
            else:
                name = 'BOM_' + str(item) + ' (' + custombom[id] + ')_v1'
            bomdic[id] = name
            d = row[-1]
            if d == True:
                d = 'T'
            elif d == False:
                d = 'F'
            else:
                d = None
            y[i] = [name, d]
            # y[i].append(row[-1])
        idic[item] = y

    # print(bomdic['a4432000000MIlL'])
    # print(idic[240160])
    return bomdic, idic

def nsitems(data):
    dic = {}
    # i = data[0].index('Item: Item Number')
    # iid = data[0].index('Product: 18 Char ID')

    i = data[0].index('Item Code')
    iid = data[0].index('Internal ID')

    for row in data[1:]:
        item = str(row[i])
        if item != None and item[:3] != 'OLD' and len(item) >= 5:
            dic[item] = row[iid]


    return dic

def dic_list_tolist(dic):
    full = []
    maxl = 0
    for key in list(dic.keys()):
        n = dic[key]
        new = [key]
        for row in n:
            for x in row:
                new.append(x)
        full.append(new)
        if len(new) > maxl:
            maxl = len(new)

    bom = maxl - 1
    bom = int(bom/2)

    head = ['Item']

    for i in range(1, bom + 1):
        x = 'Item Bills of Materials Sublist ' + str(i) + ': Bill of Materials'
        y = 'Item Bills of Materials Sublist ' + str(i) + ': Master Default'
        head.append(x)
        head.append(y)

    full.insert(0, head)

    return full

def dictolist(dic):
    full = []
    for key in list(dic.keys()):
        new = [key, dic[key]]
        full.append(new)

    return full

def BOM(bomdiclist, memo=None):
    if memo == 'Excel':
        memo = '=BOM_Memo'
    data = bomdiclist
    data[0].append('Memo')
    data[0].append('Subsidiary')
    data[0].append('Keep - Do Not Mass Delete')
    for row in data[1:]:
        row.append(memo)
        row.append(9)
        row.append('T')

    return data

def BOMrev(data, bomdic, memo=None):
    if memo == 'Excel':
        memo = '=BOM_Memo'

    head = data[0]
    item = head.index('BOM Item')
    id = head.index('BOM ID')
    head.insert(2, 'Memo')
    # head.insert(0, 'Keep - Do Not Mass Delete')
    head[item] = 'BOM Name'
    head[id] = 'BOM Revision'
    for row in data[1:]:
        bomid = row[id]
        name = bomdic[bomid]
        row[0] = name
        rev = name + '.0'
        row[1] = rev
        row.insert(2, memo)
        # row.insert(0, 'T') #For Keep do not mass delete

    return data

def customname(data):
    id = data[0].index('Identifier')
    bom = data[0].index('BOM')

    dic = {}
    for row in data[1:]:
        dic[row[bom]] = row[id]

    return dic



if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite Integration\\Items into Production 10.21.19'
    file = 'Assembly Components 12.23.19.xlsx'

    data = excel_list(file, 'Extract', folder)

    nsdic = nsitems(excel_list(file, 'Item IDs', folder))



    custBOM = customname(excel_list(file, 'Custom Names', folder))
    t = transform(data, custBOM)
    bomdic = t[0]
    itemdic = t[1]
    full = AssCompForm(data, nsdic, error=None, iid=True)


    bomlist = dictolist(bomdic)

    mem = 'Excel'
    bom = BOM(bomlist, memo=mem)
    rev = BOMrev(full, bomdic, memo=mem)

    excelcreate(dic_list_tolist(itemdic), file, 'ItemDic', folder)
    excelcreate(bom, file, 'BOM', folder)
    excelcreate(rev, file, 'BOM Revision', folder)

    # excelcreate(full, file, 'Transform', folder)

