from General import excel_list, excelcreate

def perm_format(data):
    head = data[0]

    roles = []
    permis = []

    dic = {}

    for row in data[1:]:
        role = row[head.index('Role')]
        per = row[head.index('Permission')]
        lev = row[head.index('Level')]

        if role not in roles:
            roles.append(role)

        if per not in permis:
            permis.append(per)

        dic[(role, per)] = lev

    permis.sort()
    newhead = roles.copy()
    newhead.insert(0, 'Permission')

    final = [newhead]

    for p in permis:
        nrow = [None]*len(newhead)
        nrow[0] = p

        for r in roles:
            i = newhead.index(r)
            try:
                lev = dic[(r, p)]
            except KeyError:
                lev = None

            nrow[i] = lev

        final.append(nrow)


    return final


def roles(data):

    new = [['Internal ID', 'Name', 'Role']]
    for row in data[1:]:
        if row[0] != None:
            n = row[0]
        elif row[1] != None:
            n = row[1]
        elif row[2] != None:
            n = row[2]
        elif row[3] != None:
            n = row[3]

        for x in row[5:]:
            if x != None and x != ':(':
                new.append([None, n, x])

    return new

if __name__ == '__main__':
    folder = 'C:\\Users\\jfew\\OneDrive - Allbridge\\NetSuite\\'
    # file = 'NetSuite Permissions 12.2.19.xlsx'
    #
    #
    # extract = excel_list(file, 'Extract', folder)
    #
    # data = perm_format(extract)
    #
    # print(data)
    #
    # excelcreate(data, file, 'Extract Table', folder)
    role = 'Operations Roles.xlsx'
    roledata = excel_list(role, 'Assign Roles', folder)
    newdata = roles(roledata)
    excelcreate(newdata, 'Operations Roles 12.9.19.xlsx', 'Operations', folder)
    print(newdata)