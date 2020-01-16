from General import excel_list, excelcreate
from Projects import pmapdic
import datetime as dt


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
    # file = 'Scenario 7 1.16.20 Part 1.xlsx'
    file = 'Scenario 7 1.16.20 Split.xlsx'


    multi = True
    num = 5



    if multi == True:
        comsheet = 'Sheet' + str(num)
        tsheet = 'Transform' + str(num)
    else:
        comsheet = 'Combine'
        tsheet = 'Transform'

    print('Start - ', dt.datetime.now())
    mapping = pmapdic(excel_list('SO Mapping.xlsx', 'Mapping', folder))
    print('Mapping - ', dt.datetime.now())
    data = excel_list(file, comsheet, folder)
    print('Data - ', dt.datetime.now())

    tran = transform(data, mapping)
    print('Transform - ', dt.datetime.now())

    excelcreate(tran, file, tsheet, folder)
    print('Finish - ', dt.datetime.now())