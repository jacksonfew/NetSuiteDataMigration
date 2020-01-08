from General import excel_list, excelcreate


def reformatlist(data):
    new = []
    for row in data:
        nrow =[]
        for x in row:
            if x != None and isinstance(x, int)==False:
                nrow.append(x)
        if len(nrow) != 1:
            newline = ' '.join(nrow)
        else:
            newline = nrow[0]
        newline = newline.split()
        try:
            newline = [newline[0], ' '.join(newline[1:])]
        except IndexError:
            newline = newline
        new.append(newline)

    for row in new:
        fc = row[0][0]
        punc = row[0][-1]
        if fc == 'I' or fc == 'V' or fc == 'X':
            row.append('Topic')
        elif fc in ['A', 'B', 'C', 'D', 'E', 'F']:
            row.append('Option')
            row[0] = row[0][:-1]
        else:
            if punc == ')':
                row.append('Question')
            elif punc == '.':
                row.append('Match Option')

            row[0] = int(row[0][:-1])

    return new

def createdic(data):
    full = []
    for row in data:
        t = row[-1]
        i = data.index(row)
        if t == 'Topic':
            topici = data.index(row)
        elif t == 'Question':
            qi = data.index(row)
            row.append(data[topici][1])
        else:
            data[qi].append(row)

    for row in data:
        try:
            if row[2] == 'Question':
                full.append(row)
        except IndexError:
            print(row, data.index(row))

    return full


if __name__ == '__main__':
    raw = excel_list('AdminTestQuestions.xlsx', 'Sheet1')

    reform = reformatlist(raw)
    data = createdic(reform)

    print(data)