from General import excel_list, excelcreate

def oppdic(opps):
    head = opps[0]
    id18 = head.index('18 Char ID')
    opid = head.index('Opportunity ID')
    opp = head.index('Opportunity Name')

    dic = {}

    for row in opps[1:]:
        key = (opid, opp)
        dic[key] = id18

    return dic

def addopp(data, opdic):
    data[0].append('Opp 18 ID')
    head = data[0]
    opid = head.index('Opportunity ID Ref')
    opp = head.index('Opportunity')

    full = [head]
    for row in data[1:]:
        nrow = row.copy()
        try:
            id18 = opdic[(row[opid], row[opp])]
        except KeyError:
            id18 = None

        nrow.append(id18)
        full.append(nrow)

    return full

if __name__ == '__main__':
    folder = 'C:\\Users\jfew\OneDrive - Allbridge\\Netsuite Projects\\'
    file = 'Projects SO Update.xlsx'

    data = excel_list(file, 'Case Extract', folder)

    opp = oppdic(excel_list(file, 'Opp Extract', folder))

    final = addopp(data, opp)

    excelcreate(final, file, 'Transform', folder)
