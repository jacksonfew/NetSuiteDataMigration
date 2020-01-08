from General import excel_list, excelcreate
from Projects import pmapdic


def transform(ext, mapdic):
    head = ext[0]
    newhead = list(mapdic.keys())
    full = [newhead]

    for row in ext[1:]:
        nrow = [None]*len(newhead)

        for field in newhead:
            i = newhead.index(field)
            ty = mapdic[field][0]
            loc = mapdic[field][1]
            sf = mapdic[field][2]

            if ty == 'Static':
                nrow[i] = sf
            elif ty == 'Match':
                sfi = head.index(sf)
                nrow[i] = row[sfi]
            elif ty == 'Lookup':
                if loc == 'qty':
                    qty = row[head.index('Quantity')]
                    qtyn = row[head.index('Quantity Needed')]

                    if qtyn != None:
                        nrow[i] = qtyn
                    else:
                        nrow[i] = qty

        full.append(nrow)

    return full

if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite Integration\\Sales Orders 1.6.20\\'
    file = 'Scenario 9.xlsx'

    mapping = pmapdic(excel_list('SO Mapping.xlsx', 'Mapping', folder))
    data = excel_list(file, 'Combine', folder)

    tran = transform(data, mapping)

    excelcreate(tran, file, 'Transform', folder)