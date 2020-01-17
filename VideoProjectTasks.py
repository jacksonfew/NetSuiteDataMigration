from General import excel_list, excelcreate
from GeneralProjects import template, projectname, columnsused
import openpyxl as pyxl
import os
import datetime as dt



def ms_details_dic(milestones):
    """Creates dictionary to match SF Milestone to NS Milestone name. The second dictionary stores NS milestone
    specific details (i.e. predecessor, order, and length)"""
    sfdic = {}
    nsdic = {}
    newms = {}

    for row in milestones[1:]:
        if row[1] != None:
            if row[0] not in list(sfdic.keys()):
                sfdic[row[0]] = {row[1]: row[2]}
            else:
                sfdic[row[0]][row[1]] = row[2]
        if row[1][:3] == 'New':
            if row[0] not in list(newms.keys()):
                newms[row[0]] = [row[1]]
            else:
                newms[row[0]].append(row[1])

        if row[0] not in list(nsdic.keys()):
            nsdic[row[0]] = {row[2]: row[3:]}
        else:
            nsdic[row[0]][row[2]] = row[3:]

    return sfdic, nsdic, newms


def mergemilestones(extract, mapping):

    fields = []
    for row in mapping[1:]:
        try:
            f = row[4].lower()
        except AttributeError:
            f = None
        if row[0] != 'Main' and f not in fields and f != None:
            fields.append(f)

    dic = {'Case': fields}

    sfhead = extract[0]
    for x in sfhead:
        i = sfhead.index(x)
        sfhead[i] = x.lower()


    case_i = sfhead.index('case: case number')
    for row in extract[1:]:
        for x in row:
            i = row.index(x)
            f = sfhead[i]
            if f in fields:
                f_i = fields.index(f)
                case = row[case_i]
                if case not in list(dic.keys()):
                    dic[case] = [None]*len(fields)

                if x != None and x != 0:
                    dic[case][f_i] = x

    """Right Now I have a dic with all the SF Fields needed for non main tab"""

    return dic

def addmilestones(extract, newmsdic, projlookup):
    dic = {}
    cases = []
    head = extract[0]

    name_i = head.index('Project Milestone: Project Milestone Name')
    case_i = head.index('Case: Case Number')
    crt_i = head.index('Case: Case Record Type')

    for row in extract[1:]:
        case = row[head.index('Case: Case Number')]
        crt = row[head.index('Case: Case Record Type')]
        if case not in list(dic.keys()):
            dic[case] = crt
            cases.append(case)

    for x in cases:
        # temp = template(dic[x])
        try:
            temp = projlookup[x][1]
        except KeyError:
            temp = None
        try:
            newms = newmsdic[temp]
            for m in newms:
                nrow = [None]*len(head)
                nrow[case_i] = x
                nrow[name_i] = m
                nrow[crt_i] = dic[x]

                extract.append(nrow)
        except KeyError:
            extract = extract


    return extract

def nslist(extract, map, mergedic, msdetdic, projlookup, test=False):

    head = []
    mdic = {}

    tab_i = map[0].index('Milestone')
    ns_i = map[0].index('NS Field')
    type_i = map[0].index('Type')
    loc_i = map[0].index('Where')
    sf_i = map[0].index('Field')
    var_i = map[0].index('Variable Type')

    for row in map[1:]:
        nsfield = row[ns_i]
        tab = row[tab_i]
        ty = row[type_i]
        loc = row[loc_i]
        sff = row[sf_i]
        var = row[var_i]

        head.append(nsfield)
        mdic[nsfield] = [tab, ty, loc, sff, var]

    full = [head]

    exthead = extract[0]
    for x in exthead:
        i = exthead.index(x)
        exthead[i] = x.lower()

    for row in extract[1:]:
        name = row[exthead.index('project milestone: project milestone name')]
        case = row[exthead.index('case: case number')]
        crt = row[exthead.index('case: case record type')]
        act = row[exthead.index('account name')]

        # temp = template(crt)
        try:
            temp = projlookup[case][1]
        except KeyError:
            # temp = template(crt)
            remove = True
        new = [None]*len(head)

        remove = False

        for field in head:
            i = head.index(field)
            tab = mdic[field][0]
            ty = mdic[field][1]
            loc = mdic[field][2]
            sff = mdic[field][3]

            if tab == 'Main':
                if ty == 'Static' or ty == 'dependent static':
                    new[i] = sff
                elif ty == 'SF':
                    sf_i = exthead.index(sff.lower())
                    new[i] = row[sf_i]
                elif ty == 'lookup':
                    if loc == 'Milestone Name':
                        try:
                            new[i] = msdetdic[0][temp][name]
                        except:
                            # remove = True
                            new[i] = name
                    elif loc == 'Project':
                        # new[i] = projectname(case, act, crt, test)
                        try:
                            new[i] = projlookup[case][0]
                        except KeyError:
                            remove = True
                    elif loc == 'Milestone Order':
                        try:
                            new[i] = msdetdic[1][temp][msdetdic[0][temp][name]][1]
                        except:
                            new[i] = None

                    elif loc == 'Order':
                        try:
                            new[i] = msdetdic[1][temp][msdetdic[0][temp][name]][5]
                        except:
                            # remove = True
                            order = row[exthead.index('milestone order')]
                            if order != None:
                                new[i] = order
                            else:
                                remove = True
                    elif loc == 'predessor':
                        try:
                            new[i] = msdetdic[1][temp][msdetdic[0][temp][name]][0]
                        except:
                            new[i] = None
                    elif loc == 'length':
                        try:
                            new[i] = msdetdic[1][temp][msdetdic[0][temp][name]][2]
                        except:
                            new[i] = None
                    elif loc == 'Default Resource':
                        try:
                            new[i] = msdetdic[1][temp][msdetdic[0][temp][name]][3]
                        except:
                            new[i] = None
                    elif loc == 'Subtab':
                        try:
                            new[i] = msdetdic[1][temp][msdetdic[0][temp][name]][4]
                        except:
                            new[i] = None
                    elif loc == 'Delete':
                        if test == True:
                            new[i] = 'F'
                        else:
                            new[i] = 'T'

            else:
                if ty == 'SF':
                    values = mergedic[case]
                    v_i = mergedic['Case'].index(sff.lower())
                    new[i] = values[v_i]

        # full.append(new)
        if remove == False:
            full.append(new)

    full = variableconvert(mdic, full)

    for row in full[1:]:
        r_i = full.index(row)
        task = row[head.index('Name')]
        if row[head.index('PJ Task Subtab Link')] == None:
            legacy = True
        else:
            legacy = False
        for f in row:
            i = row.index(f)
            fname = head[i]
            fdic = mdic[fname]
            ftab = fdic[0]
            fty = fdic[1]
            floc = fdic[2]

            if fty == 'dependent static':
                loci = head.index(floc)
                if row[loci] == None:
                    full[r_i][i] = None

            if ftab != 'Main' and ftab != task and legacy != True:
                full[r_i][i] = None


    sortdata = full[1:]
    sortdata.sort(reverse=True)

    final = [full[0]]
    for row in sortdata:
        final.append(row)


    return final


def variableconvert(mapdic, data):
    """Designing to be used within nslist function"""

    head = data[0]

    for row in data[1:]:
        ri = data.index(row)
        for f in row:
            i = row.index(f)
            var = mapdic[head[i]][-1]
            if f != None:
                if var == 'Date':
                    if isinstance(f, dt.datetime):
                        new = dt.date(f.year, f.month, f.day)
                    elif isinstance(f, dt.date):
                        new = f
                    else:
                        new = dt.date(2019, 7, 4)
                elif var == 'Free-Form Text':
                    new = str(f)
                elif var == 'Integer Number':
                    new = int(f)
                else:
                    new = f
            else:
                new = None

            data[ri][i] = new

    return data

def projectlookup(projectextract):
    '''THIS IS THE OLD VERSION'''
    head = projectextract[0]
    data = projectextract[1:]
    id = head.index('External ID')
    case = head.index('Case Number')

    dic = {}
    for row in data:
        dic[row[case]] = row[id]

    return dic

def projectInternalID(data, test=False):
    dic = {}
    for row in data:
        case = row[1]
        iid = row[0]
        name = row[2]
        temp = row[3]

        if test == False:
            dic[case] = [iid, temp]
        else:
            dic[case] = [name, temp]
    return dic


def missingPred(data):
    head = data[0]
    casei = head.index('Case (JF Only)')
    msi = head.index('Name')
    predi = head.index('Predecessor Sublist: Task')
    predconi = head.index('Predecessor Sublist: Type')

    dic = {}
    for row in data[1:]:
        case = row[casei]
        ms = row[msi]

        if case not in list(dic.keys()):
            dic[case] = [ms]
        else:
            dic[case].append(ms)

    for row in data[1:]:
        i = data.index(row)
        case = row[casei]
        ms = row[msi]
        pred = row[predi]

        tasks = dic[case]
        if pred not in tasks:
            data[i][predi] = None
            data[i][predconi] = None

    return data

def owner(data, owndata):
    dic = {}
    for row in owndata:
        dic[row[0]] = row[1]

    head = data[0]
    for row in data[1:]:
        owner = row[head.index('Milestone Owner')]
        dr = row[head.index('Default Resource')]

        try:
            name = dic[owner]
        except KeyError:
            if dr != None:
                name = dr
            else:
                name = 'Project Manager'

        row[head.index('Milestone Owner')] = name
        row[head.index('Default Resource')] = name

    return data

if __name__ == '__main__':
    folder = 'C:\\Users\jfew\OneDrive - Allbridge\\Netsuite Projects\\'
    mapping = 'Project and Milestone Mapping.xlsx'
    # extract = 'Extract 12.31.19.xlsx'
    extract = 'Extract 1.17.20.xlsx'
    writefile = 'Milestone Transform.xlsx'

    data = excel_list(extract, 'Milestones', folder)

    ms_data = excel_list(mapping, 'Milestones', folder)
    ms_map = excel_list(mapping, 'Milestone Fields', folder)
    milestonedics = ms_details_dic(ms_data)
    newmsdic = milestonedics[2]
    print(milestonedics[0])
    print(milestonedics[1]['Headend']['PTR'])

    test = False


    projlookup = projectInternalID(excel_list('NetSuite Projects.xlsx', 'Projects', folder), test)
    data = addmilestones(data, newmsdic, projlookup)





    mergedic = mergemilestones(data, ms_map)
    # projlookup = projectlookup(excel_list(extract, 'Projects', folder))

    # print(mergedic)
    newdata = nslist(data, ms_map, mergedic, milestonedics, projlookup, test)

    newdata = missingPred(newdata)

    newdata = owner(newdata, excel_list('NetSuite Projects.xlsx', 'Owner', folder))

    excelcreate(newdata, writefile, 'All Milestones', folder)

"""COLUMNSUSED IS NOT WORKING PROPERLY KMS"""
"""NS Mapping is actually pretty automated so this really isnt needed"""
    # columns = columnsused(newdata)
    # excelcreate(columns, writefile, 'Milestones', folder)
    # print(len(columns[0]))
