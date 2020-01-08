from General import excel_list, excelcreate
from GeneralProjects import template, projectname, dependmod, systemtemp
import openpyxl as pyxl
import os
import datetime as dt


def pmapdic(map):
    dic = {}
    for row in map[1:]:
        if row[1] == 'Combo':
            dic[row[0]] = [row[1], row[2], row[3].split(', ')]
        else:
            dic[row[0]] = row[1:]
    return dic


def customerlookup():
    folder = 'C:\\Users\jfew\OneDrive - Allbridge\\Netsuite Projects\\'
    nsextract = excel_list('Lookups for Projects.xlsx', 'Customers', folder)
    head = nsextract[0]
    internal = head.index('Internal ID')
    custnum = head.index('SF Customer Number')
    account = head.index('Name')
    dic = {}
    numv = None
    for row in nsextract[1:]:
        id = row[internal]
        num = row[custnum]
        act = row[account]

        if num != None:
            dic[num] = id
        if act != None:
            dic[act] = id

    # if cnum != None:
    #     try:
    #         numv = dic[cnum]
    #     except KeyError:
    #         numv = None
    #
    # try:
    #     aval = dic[aname]
    # except KeyError:
    #     aval = None
    #
    # if numv != None:
    #     ans = numv
    # else:
    #     if aval != None:
    #         ans = aval
    #     else:
    #         ans = error
    #
    # if test == True:
    #     ans = 16163

    return dic

def projectSFIDdic():
    folder = 'C:\\Users\jfew\OneDrive - Allbridge\\Netsuite Projects\\'
    idlist = excel_list('Lookups for Projects.xlsx', 'External ID', folder)

    dic = {}

    for row in idlist[1:]:
        proj = row[0]
        id = row[1]
        dic[proj] = id

    return dic


def ptrans(ext, mapdic, test=False):

    head = ext[0]
    newhead = list(mapdic.keys())
    full = [newhead]
    projectlookup = projectSFIDdic()
    cuslookup = customerlookup()
    print(len(ext[1:]))
    num = 0
    for row in ext[1:]:
        new = [None]*len(newhead)
        for field in newhead:
            i = newhead.index(field)
            ty = mapdic[field][0]
            loc = mapdic[field][1]
            sf = mapdic[field][2]
            if ty[-6:].lower() == 'static':
                new[i] = sf
            elif ty == 'SF':
                sfi = head.index(sf)
                new[i] = row[sfi]
            elif ty == 'Combo':
                com = [None]*len(sf)
                for x in sf:
                    xi = sf.index(x)
                    sfi = head.index(x)
                    com[xi] = str(row[sfi])
                case = com[sf.index('Case Number')]
                act = com[sf.index('Account Name')]
                crt = com[sf.index('Case Record Type')]
                ost = com[sf.index('Opp System Type')]
                t = com[sf.index('Type')]
                if loc == 'project function':
                    new[i] = projectname(case, act, crt, ost, t, test)[0]
                elif loc == 'description':
                    new[i] = projectname(case, act, crt, ost, t, test)[1]
                elif loc == 'template':
                    new[i] = systemtemp(crt, ost, t)[0]
                elif loc == 'system type':
                    new[i] = systemtemp(crt, ost, t)[1]

            elif ty == 'lookup':
                if loc == 'template':
                    sfi = head.index(sf)
                    new[i] = template(row[sfi])
                elif loc == 'Customer':
                    act = row[head.index('Account Name')]
                    cnum = row[head.index('Customer Number')]
                    # new[i] = customerlookup(cnum, act, test)
                    try:
                        new[i] = cuslookup[cnum]
                    except KeyError:
                        try:
                            new[i] = cuslookup[act]
                        except KeyError:
                            new[i] = 'Error'
                elif loc == 'externalid':
                    case = row[head.index('Case Number')]
                    try:
                        new[i] = projectlookup[case]
                    except KeyError:
                        new[i] = 'Missing'
                elif loc == 'Delete':
                    if test == True:
                        new[i] = 'F'
                    else:
                        new[i] = 'T'
            elif ty == 'if blank':
                sfi = head.index(sf)
                v = row[sfi]
                if v == None:
                    new[i] = loc
                else:
                    new[i] = v
            elif ty == 'T/F':
                sfi = head.index(sf)
                v = row[sfi]
                if v == 1:
                    new[i] = 'T'
                else:
                    new[i] = 'F'
        num += 1
        print(num)
        full.append(new)

    for row in full:
        ri = full.index(row)
        for f in row:
            fi = row.index(f)
            if isinstance(f, int):
                n = str(f)
            elif isinstance(f, dt.datetime):
                n = dt.date(f.year, f.month, f.day)
            else:
                n = f

            full[ri][fi] = n

    full = dependmod(full, mapdic)

    return full





if __name__ == '__main__':
    folder = 'C:\\Users\jfew\OneDrive - Allbridge\\Netsuite Projects\\'
    mapping = 'Project and Milestone Mapping.xlsx'
    extract = 'Extract 12.31.19.xlsx'
    writefile = 'Milestone Transform.xlsx'
    writesheet = 'Projects'

    test = False


    mapdic = pmapdic(excel_list(mapping, 'Project Fields', folder))
    ext = excel_list(extract, 'Projects', folder)

    final = ptrans(ext, mapdic, test)
    # print(mapdic)
    print(final)

    excelcreate(final, writefile, writesheet, folder)
    # print(customerlookup('C67375', None, 'Missing'))
