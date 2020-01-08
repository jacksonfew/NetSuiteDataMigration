from General import excel_list, excelcreate





def gla_dic(gla_list):
    dic = {}
    dic['Headers'] = gla_list[0]
    for row in gla_list[1:]:
        dic[row[0]] = row

    return dic

def extid_dic(extid):
    dic = {}
    for row in extid[1:]:
        dic[row[0]] = row[1]

    return dic

def unit_dic(units):
    dic = {}
    # dic['Headers'] = units[0]
    for row in units[1:]:
        dic[row[0]] = row[2:]

    return dic

def binloc(list):
    dic = {}
    for row in list:
        sf = row[0]
        dic[sf] = [row[1], row[2], row[4], row[5]]

    return dic

def vendors(ns):
    dic = {}
    name = ns[0].index('Name')
    id = ns[0].index('Internal ID')
    for row in ns[1:]:
        dic[row[name]] = row[id]

    return dic



def ns_list(guide, extract, gla, extid, units, bins, vend, error=None, iID=True):
    headers = []
    h_dic = {}
    for row in guide[1:]:
        headers.append(row[0])
        h_dic[row[0]] = row[1:]
    full = [headers]

    sf_head = extract[0]
    for row in extract[1:]:
        nrow = len(headers)*[None]
        item = row[sf_head.index('Item: Item Number')]
        unit = row[sf_head.index('Default Unit of Measure')].lower()

        for field in headers:
            i = headers.index(field)
            r = h_dic[field]

            if r[0] == 'Static':
                nrow[i] = r[2]
            elif r[0] == 'SF':
                sf_i = sf_head.index(r[2])
                if r[2] == 'Maximum On Hand' or r[2] == 'Minimum On Hand':
                    if row[sf_i] == -1:
                        row[sf_i] = 0

                nrow[i] = row[sf_i]
            elif r[0] == 'lookup':
                if r[1] == 'GLA Mod':
                    h = gla['Headers']
                    if iID == True:
                        h_i = h.index(field)
                    else:
                        h_i = h.index(field+' Name')
                    try:
                        nrow[i] = gla[item][h_i]
                    except KeyError:
                        nrow[i] = gla['Default'][h_i]
                elif r[1] == 'External ID':
                    if item in list(extid.keys()):
                        nrow[i] = extid[item]
                elif r[1] == 'Units':
                    try:
                        val = units[unit]
                    except KeyError:
                        val = units['each (ea)']
                    if iID == False:
                        nrow[i] = unit
                    else:
                        if field == 'Units Type':
                            nrow[i] = val[0]
                        else:
                            nrow[i] = val[1]
                elif r[1] == 'Location' or r[1] == 'Bin':
                    sf_i = sf_head.index(r[2])
                    val = row[sf_i]
                    if val in list(bins.keys()):
                        if r[1] == 'Location':
                            ans = [bins[val][2], bins[val][3]]
                        else:
                            ans = [val, bins[val][1]]
                    else:
                        ans = [val, error]

                    if iID == True:
                        nrow[i] = ans[1]
                    else:
                        nrow[i] = ans[0]
                elif r[1] == 'Vendor':
                    sf_i = sf_head.index(r[2])
                    val = row[sf_i]
                    try:
                        ans = vend[val]
                    except KeyError:
                        ans = error

                    if iID == False:
                        ans = val

                    nrow[i] = ans
            elif r[0] == 'Convert':
                sf_i = sf_head.index(r[2])
                val = row[sf_i]
                if r[1] == 'T/F':
                    if val == 0 or val == '0':
                        val = 'F'
                    else:
                        val = 'T'

                nrow[i] = val

        full.append(nrow)

    return full

def assItemdic(sheetdata):
    sheetdata[0][0] = 'Headers'

    dic = {}
    for row in sheetdata:
        dic[row[0]] = row[1:]

    return dic





def addAss(data, itemdic):
    head = data[0]
    i = head.index('Item Code')

    for n in itemdic['Headers']:
        head.append(n)

    for row in data[1:]:
        try:
            x = itemdic[row[i]]
        except KeyError:
            x = False

        if x != False:
            for field in x:
                row.append(field)

    return data



if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite Integration\\Items into Production 10.21.19'
    extractfile = 'Extract 1.2.20.xlsx'
    mapfile = 'Item Mapping.xlsx'
    assfile = 'Assembly Components 12.23.19.xlsx'

    itemtypes = ['Regular', 'Serial', 'Assembly', 'Serial Assembly']

    savefiles = {'Regular':'Regular Items 10.21.19.xlsx', 'Serial':'Serial Items 10.22.19.xlsx', \
                'Assembly':'Non-Serial Assembly 10.22.19.xlsx', 'Serial Assembly':'Serial Assembly 10.23.19.xlsx'}

    print('Please Select Item Type')
    for x in itemtypes:
        n = str(itemtypes.index(x)) + ': ' + x
        print(n)


    i = int(input())
    item = itemtypes[i]


    file = savefiles[item]
    print(file)

    gla = gla_dic(excel_list('GLA MOD.xlsx', 'Mod', folder))
    extid = extid_dic(excel_list('External_ID.xlsx', 'Sheet1', folder))
    units = unit_dic(excel_list('Units of Measure 10.22.19.xlsx', 'Summary', folder))
    bins = binloc(excel_list('Bins and Locations 10.22.19.xlsx', 'Summary', folder))
    ven = vendors(excel_list('Vendors 10.22.19.xlsx', 'Vendors', folder))
    itemdic = assItemdic(excel_list(assfile, 'ItemDic', folder))






    guide = excel_list(mapfile, item, folder)
    extract = excel_list(extractfile, item, folder)

    full = ns_list(guide, extract, gla, extid, units, bins, ven, error=None, iID=True)

    if item.find('Assembly') != -1:
        full = addAss(full, itemdic)

    excelcreate(full, file, 'Transform', folder)

