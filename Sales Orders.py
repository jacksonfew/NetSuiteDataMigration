from General import excel_list, excelcreate


def casedic(data):
    head = data[0]
    so_i = head.index('Sales Order: Sales Order #')
    c1_id_i = head.index('Install Case ID')
    c1_i = head.index('Install Case')
    c2_id_i = head.index('Related Case ID')
    c2_i = head.index('Related Case')

    dic = {}
    for row in data[1:]:
        so = row[so_i]
        if so not in list(dic.keys()):
            case1 = row[c1_i]
            case2 = row[c2_i]
            if case1 != None:
                dic[so] = [row[c1_id_i], case1]
            else:
                dic[so] = [row[c2_id_i], case2]

    return dic

def productdic(data):
    head = data[0]
    so_i = head.index('Sales Order: Sales Order #')
    status_i = head.index('Status')
    prodID_i = head.index('Product 18 Char ID')
    item_i = head.index('Item Number')
    des_i = head.index('Description')
    qty_i = head.index('Quantity Needed')
    pack_i = head.index('Quantity Packed')

    dic = {}

    for row in data:
        key = (row[so_i], row[item_i])
        dic[key] = [row[status_i], row[prodID_i], row[item_i], row[des_i], row[qty_i], row[pack_i]]

    return dic


def sostatus(data):
    head = data[0]
    so_i = head.index('Sales Order: Sales Order #')
    status_i = head.index('Status')

    dic = {}

    for row in data[1:]:
        so = row[so_i]
        status = row[status_i]

        if so not in list(dic.keys()):
            dic[so] = [status]
        else:
            if status not in dic[so]:
                dic[so].append(status)

    return dic


def combine(data, cdic, pdic, sodic):
    head = data[0]
    so_i = head.index('Sales Order')
    cid = head.index('Case ID')
    c = head.index('Case')
    s = head.index('Status')
    qty_i = head.index('Quantity Needed')
    # pack_i = head.index('Quantity Packed')
    # prodID_i = head.index('18 Char ID')
    item_i = head.index('Product Code')
    des_i = head.index('Product Name')

    full = [head]
    for row in data[1:]:
        nrow = row.copy()
        so = row[so_i]
        try:
            case = cdic[so]
        except KeyError:
            case = [None, None]

        nrow[cid] = case[0]
        nrow[c] = case[1]

        try:
            item = pdic[(so, row[item_i])]
        except KeyError:
            item = [None]*6

        nrow[s] = item[0]
        nrow[qty_i] = item[4]
        # nrow[pack_i] = item[5]

        if nrow[s] == None:
            try:
                nrow[s] = sodic[so][0]
            except KeyError:
                nrow[s] = 'Open'

        if nrow[s] != 'Cancelled':
            full.append(nrow)

    return full



if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite Integration\\Sales Orders 1.6.20\\'
    file = 'Scenario 8.xlsx'

    sfdata = excel_list('Ascent SO Lines Scenario 8.xlsx', 'Ascent SO', folder)
    kvdata = excel_list(file, 'Master', folder)

    pdic = productdic(sfdata)
    cdic = casedic(sfdata)
    sodic = sostatus(sfdata)

    final = combine(kvdata, cdic, pdic, sodic)

    excelcreate(final, file, 'Combine', folder)