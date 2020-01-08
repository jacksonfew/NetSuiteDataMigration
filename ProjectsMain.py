from General import excel_list, excelcreate
from GeneralProjects import template, projectname, columnsused
import VideoProjectTasks as tasks
import Projects as proj


folder = 'C:\\Users\jfew\OneDrive - Allbridge\\Netsuite Projects\\'
mapfile = 'Project and Milestone Mapping.xlsx'
extractfile = 'Extract 11.22.19.xlsx'
writefile = 'Projects Transform.xlsx'


def projectsfinal(data, test=False):
    mapdic = proj.pmapdic(excel_list(mapfile, 'Project Fields', folder))
    final = proj.ptrans(data, mapdic, test)

    return final

def tasksfinal(data, test=False):

    ms_data = excel_list(mapfile, 'Milestones', folder)
    ms_map = excel_list(mapfile, 'Milestone Fields', folder)
    milestonedics = tasks.ms_details_dic(ms_data)
    newmsdic = milestonedics[2]

    data = tasks.addmilestones(data, newmsdic)
    mergedic = tasks.mergemilestones(data, ms_map)
    newdata = tasks.nslist(data, ms_map, mergedic, milestonedics, test)
    # final = columnsused(newdata)
    final = newdata

    return final

def mainoptions():
    data = excel_list(mapfile, 'Data Selection', folder)
    head = data[0]
    dic = {}
    for row in data[1:]:
        crt = row[0]
        for n in range(0, len(head)):
            cat = head[n]

            if row[n] == True:
                if cat not in list(dic.keys()):
                    dic[cat] = [crt]
                else:
                    dic[cat].append(crt)

    return dic

def optionselect():
    optdic = mainoptions()
    opt = list(optdic.keys())
    print('Import Data Options:')
    for x in opt:
        i = opt.index(x)
        print(str(i) + ' : ' + x)

    sel = opt[int(input('Please Select an Option: '))]

    if sel != 'Individual':
        crt = optdic[sel]
    else:
        crt = []
        done = False
        l = optdic[sel]
        print('Case Record Types:')
        for x in l:
            i = l.index(x)
            print(str(i) + ' : ' + x)

        y = int(input('Select CRT: '))
        crt.append(l[y])

    return crt

def removedata(projfull, taskfull, select):
    projhead = projfull[0]
    taskhead = taskfull[0]

    pdata = [projhead]
    tdata = [taskhead]

    pi = projhead.index('Case Record Type')
    ti = taskhead.index('Case: Case Record Type')
    for row in projfull[1:]:
        if row[pi] in select:
            pdata.append(row)

    for row in taskfull[1:]:
        if row[ti] in select:
            tdata.append(row)

    return pdata, tdata




if __name__ == '__main__':

    test = True

    projext = excel_list(extractfile, 'Projects', folder)
    taskext = excel_list(extractfile, 'Milestones', folder)
    select = optionselect()
    re = removedata(projext, taskext, select)

    project = projectsfinal(re[0], test)
    task = tasksfinal(re[1], test)

    excelcreate(project, writefile, 'Projects', folder)
    excelcreate(task, writefile, 'Milestones', folder)
