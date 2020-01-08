from General import excel_list, excelcreate, dateconvert, project, Taskdetail, task_len
import datetime as dt



def task_transform(data, mapping):
    datahead = data[0]
    newhead = []
    for row in mapping[1:]:
        newhead.append(row[0])
    datanew = [newhead]

    datahead_dic = {}
    for field in datahead:
        datahead_dic[field] = datahead.index(field)

    special = ['Name', 'Project Name', 'Status', 'Insert Before', 'Constraint Type', 'Predecessor Sublist: Task', \
               'Predecessor Sublist: Type', 'Task Length in Days', 'PJ Task Subtab Link']

    for row in data[1:]:
        sdic = {}
        newrow = [None]*len(newhead)

        old_tname = row[datahead.index('Project Milestone: Project Milestone Name')]
        old_pretask = row[datahead.index('Predecessor Name')]
        old_crt = row[datahead.index('Case: Case Record Type')]
        case = row[datahead.index('Case: Case Number')]
        act = row[datahead.index('Account Name')]
        tstart = row[datahead.index('Ready to Start')]
        tend = row[datahead.index('Completed On')]

        try:
            tname = Taskdetail()[old_crt][old_tname][0]
        except KeyError:
            tname = None

        sdic['Name'] = tname
        sdic['Project Name'] = project(case, act, old_crt)
        sdic['Status'] = 'Not Started'
        sdic['Insert Before'] = Taskdetail()[old_crt][old_tname][1]
        sdic['Constraint Type'] = None                              #Not Sure
        try:
            sdic['Predecessor Sublist: Task'] = Taskdetail()[old_crt][old_pretask][0]
        except KeyError:
            sdic['Predecessor Sublist: Task'] = None
        sdic['Predecessor Sublist: Type'] = None                    #Not Sure
        sdic['Task Length in Days'] = task_len(tstart, tend)
        sdic['PJ Task Subtab Link'] = tname

        for r in mapping[1:]:
            field = r[0]
            i = newhead.index(field)
            include = r[mapping[0].index(tname)]
            if include == True:
                if field in special:
                    newrow[i] = sdic[field]
                else:
                    sf = r[1]
                    sf_i = datahead_dic[sf]
                    newrow[i] = row[sf_i]
        datanew.append(newrow)
    return datanew


if __name__ == '__main__':
    tmap = excel_list('Milestone Field Mapping.xlsx', 'COMM')
    data = excel_list('COMM Task Test 10.16.19.xlsx', 'Sheet1')

    newdata = dateconvert(task_transform(data, tmap))

    excelcreate(newdata,'COMM Task Test 10.16.19.xlsx', 'Sheet2')
