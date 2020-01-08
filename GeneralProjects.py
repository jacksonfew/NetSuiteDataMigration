from General import excel_list, excelcreate

def template(crt):
    if crt == 'HD Install' or crt == 'SD Install':
        ans = 'Headend'
    elif crt == 'DRE Install':
        ans = 'DRE'
    elif crt == 'COMM Install' or crt == 'COMM HE Install':
        ans = 'Commercial'
    elif crt == 'TV Set Install':
        ans = 'TV Set Install'
    elif crt == 'Internet Install' or crt == 'Mitel Install':
        ans = 'HSIA/Voice'
    elif crt[-2:] == 'SS':
        ans = 'Site Survey'
    else:
        ans = 'NO UPDATE: ' + crt

    return ans

def r(cur, new):
    if cur == None:
        ans = new
    else:
        ans = cur
    return ans

def systemtemp(crt, ost, type):
    t = None
    st = None
    if crt[-2:] == 'SS':
        t = 'Site Survey'
    if type == 'IHG Studio':
        st = type
    if crt == 'Cabling Install' or ost == 'Cabling Only':
        t = r(t, 'Cabling Install')
        st = r(st, 'Cabling')
    if type == 'Allbridge Entertainment':
        t = r(t, type)
        st = r(st, type)
    if crt == 'HD Install':
        t = r(t, 'Headend')
        if type == 'Receiverless HD':
            st = r(st, 'Receiverless HD Headend')
        else:
            st = r(st, 'HD Headend')
    if crt == 'DRE Install' or crt =='DRE SS':
        t = r(t, 'DRE')
        if ost == 'DRE Plus':
            st = r(st, 'ADV DRE')
        else:
            st = r(st, 'Basic DRE')
    if crt == 'Headend SS':
        st = r(st, 'Video')
    if crt == 'SD Install':
        if type == 'D12 Swap Install':
            t = r(t, 'D12 - Eng')
        else:
            t = r(t, 'Headend')

        st = r(st, 'SD Headend')
    if crt == 'TV Set Install':
        t = r(t, 'TV Set Install')
        st = r(st, 'Panels')
    if type == 'Faceplate Only':
        t = r(t, 'Faceplates')
    if crt == 'Mitel Install' or crt == 'Mitel SS':
        t = r(t, 'HSIA/Voice')
        st = r(st, 'Voice')
    if crt == 'Internet Install' or crt == 'Internet SS':
        t = r(t, 'HSIA/Voice')
        st = r(st, 'Data')
    if crt == 'Network SS':
        st = r(st, 'Network')

    if crt[:6] == 'Hilton':

        t = r(t, 'Hilton CR')
        st = r(st, 'IOT - Hilton Connected Room')
    if crt[:4] == 'COMM':
        t = r(t, 'Commercial')
        st = r(st, 'Video')
    return t, st

def systemtypeTEST(data):
    head = data[0]
    full = [head]
    for row in data[1:]:
        nrow = row[:3]
        func = systemtemp(row[0], row[1], row[2])
        nrow.append(func[1])
        nrow.append(func[0])
        full.append(nrow)
    return full

def projectname(case, act, crt, ost, type, test=False):
    if test == True:
        act = 'Data Migration Final Tester'
    info = systemtemp(crt, ost, type)
    temp = info[0]
    st = info[1]
    if temp == 'Site Survey':
        t = st + ' ' + 'SS'
    else:
        t = st
    des = str(act) + ' ' + str(t)
    name = str(case) + ' ' + des
    return name, des

def dependmod(full, mapdic):
    """Designed to be used in transform function to deal with dependent fields"""

    head = full[0]

    for row in full[1:]:
        ri = full.index(row)
        for f in row:
            fi = row.index(f)
            fname = head[fi]
            ty = mapdic[fname][0]
            loc = mapdic[fname][1]
            sf = mapdic[fname][2]
            if ty != None:
                if ty[:9].lower() == 'dependent':
                    loci = head.index(loc)
                    if row[loci] == None:
                        full[ri][fi] = None
    return full

def columnsused(full):
    """Eliminates all fields that don't have any values for any row. Done because Mapping in NetSuite has to be done
    every load. Thus this cuts down on the number of fields that must be mapped"""
    dic = {}
    head = full[0]
    for x in head:
        dic[x] = 0

    for row in full[1:]:
        for x in row:
            i = row.index(x)
            field = head[i]
            if x != None:
                dic[field] += 1

    newhead = []
    for f in head:
        if dic[f] != 0:
            newhead.append(f)

    newdata = [newhead]

    for row in full[1:]:
        newrow = [None] * len(newhead)
        for x in row:
            i = row.index(x)
            field = head[i]

            if dic[field] != 0:
                newi = newhead.index(field)
                newrow[newi] = x
        newdata.append(newrow)

    return newdata


if __name__ == '__main__':
    folder = 'C:\\Users\jfew\OneDrive - Allbridge\\Netsuite Projects\\'
    file = 'System Types.xlsx'

    data = excel_list(file, 'Sheet4', folder)

    moddata = systemtypeTEST(data)

    excelcreate(moddata, file, 'Sheet4', folder)