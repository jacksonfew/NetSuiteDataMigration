from General import excel_list, excelcreate, unit_dic, binloc
from AssemblyComponents import nsitems
import datetime as dt



def Adj(extract, guide, items, bins, units):
    headers = []
    h_dic = {}
    for row in guide[1:]:
        headers.append(row[0])
        h_dic[row[0]] = row[1:]
    full = [headers]
    sf_head = extract[0]

    for row in extract[1:]:
        nrow = len(headers)*[None]

        for field in headers:
            i = headers.index(field)
            r = h_dic[field]
            type = r[0]
            where = r[1]
            field = r[2]

            if type == 'Static':
                nrow[i] = field
            elif type == 'SF':
                sf_i = sf_head.index(field)
                nrow[i] = row[sf_i]
            elif type == 'lookup':
                if where == 'Units':
                    sf_i = sf_head.index(field)
                    unit = row[sf_i]
                    if unit == None:
                        unit = 'each'
                    try:
                        unit = units[unit.lower()]
                    except KeyError:
                        unit = [1, 1]
                    nrow[i] = unit[1]
                elif where == 'Item':
                    sf_i = sf_head.index(field)
                    val = str(row[sf_i])
                    try:
                        nrow[i] = items[val]
                    except KeyError:
                        nrow[i] = None
                elif where == 'Location' or where == 'Bin':
                    sf_i = sf_head.index(field)
                    val = row[sf_i]

                    if where == 'Location':
                        try:
                            bin = bins[val][3]
                        except KeyError:
                            bin = 1
                    else:
                        try:
                            bin = bins[val][1]
                        except KeyError:
                            bin = None

                    nrow[i] = bin

        full.append(nrow)

    nobin = [headers]
    yesbin = [headers]
    bini = headers.index('Adjustment Sublist: Bin')

    for row in full[1:]:
        if row[bini] == None:
            nobin.append(row)
        else:
            yesbin.append(row)

    return full, yesbin, nobin


def finaladj(data, memo):
    head = data[0]
    ei = head.index('External ID')
    mi = head.index('Memo')
    lid = head.index('Adjustment Sublist: Line ID')

    d = dt.datetime.now()
    extid = 'JF' + str(d.year) + str(d.month) + str(d.day) + str(d.hour) + str(d.minute) +  \
            str(d.second) + str(d.microsecond)

    for row in data[1:]:
        i = data.index(row)
        row[lid] = 'B' + str(i)
        row[ei] = extid
        row[mi] = memo

    return data



if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite Integration\\Items into Production 10.21.19'
    nsfile = 'Netsuite Items.xlsx'
    extractfile = 'Adjustement Extract 1.2.20.xlsx'
    mapfile = 'Adjustment Mapping.xlsx'
    transfile = 'Adjustment Transform.xlsx'
    adjtype = ['Regular', 'Serial']
    sheetdic = {'Regular': ['Bin', 'NoBin'], 'Serial': ['BinSerial', 'NoBinSerial']}


    print('Please Select an Item Type')
    for x in adjtype:
        msg = str(adjtype.index(x)) + ': ' + x
        print(msg)
    i = int(input())
    selec = adjtype[i]


    units = unit_dic(excel_list('Units of Measure 10.22.19.xlsx', 'Summary', folder))
    bins = binloc(excel_list('Bins and Locations 10.22.19.xlsx', 'Summary', folder))
    items = nsitems(excel_list(nsfile, 'Items', folder))




    extract = excel_list(extractfile, selec, folder)
    guide = excel_list(mapfile, selec, folder)

    adj = Adj(extract, guide, items, bins, units)
    # full = adj[0]
    yesbin = finaladj(adj[1], None)
    nobin = finaladj(adj[2], None)

    # excelcreate(full, transfile, 'Full', folder)
    excelcreate(yesbin, transfile, sheetdic[selec][0], folder)
    excelcreate(nobin, transfile, sheetdic[selec][1], folder)
    # print(full)