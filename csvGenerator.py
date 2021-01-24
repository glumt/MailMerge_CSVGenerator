import csv
import os

def readCSV(filepath, dlm = ';', noHeader = 1):
    with open(filepath, encoding = 'utf-8') as csvfile:
        readCSV = csv.reader(csvfile, delimiter = dlm)
        header = []
        table = []
        for i, row in enumerate(readCSV):
            if i < noHeader:
                header = row
            else:
                table.append(row)
    return header, table

def main():
    csvInput = 'mail_temp.csv'
    header, table = readCSV(csvInput)

    # strip all surrounding spaces
    header = [h.strip() for h in header]
    table = [[v.strip() for v in row] for row in table]

    # key to list id mapping
    h2id = {d.strip(): i for i, d in enumerate(header)}

    # get attachments from folders
    attachments = []
    for row in table:
        firstName = row[h2id["Vorname"]]
        lastName = row[h2id["Nachname"]]
        dirPath = f'{firstName} {lastName}'
        onlyfiles = [f for f in os.listdir(dirPath) if os.path.isfile(os.path.join(dirPath, f))]

        fullpaths = [os.path.abspath(os.path.join(dirPath, f)) for f in onlyfiles]
        attachments.append(fullpaths)
    noMaxAtt = max([len(r) for r in attachments])

    # fill empty attachment entries
    for i, row in enumerate(attachments):
        if len(row) < noMaxAtt:
            emptyEntries = ['""' for i in range(noMaxAtt - len(row))]
            row = row + emptyEntries
            attachments[i] = row

    # create new header line
    newHeader = header + [f"Anhang{i+1}" for i in range(noMaxAtt)]

    # create new csv file with file attachments
    with open('mailData.csv', 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter = ';', quotechar='\'')#, quoting=csv.QUOTE_NONE)
        csvwriter.writerow(newHeader)
        for row, att in zip(table, attachments):
            csvwriter.writerow(row + att)

    # remove last newline from file
    with open('mailData.csv', 'r', newline='', encoding='utf-8') as f:
        filelines = f.readlines()
        newLastLine = filelines[-1].replace(os.linesep, '')
        filelines[-1] = newLastLine
    with open('mailData.csv', 'w', newline='', encoding='utf-8') as f:
        for row in filelines:
            print(row)
            f.write(row)

    return 0

if __name__ == '__main__':
    main()
