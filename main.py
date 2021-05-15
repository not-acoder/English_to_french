import xlrd
d = {}

loc = (r"C:\Users\Abhinav\Desktop\french_dictionary.xls")
fin_loc=(r"C:\Users\Abhinav\Desktop\t8.shakespeare.txt")
fout_loc=(r"C:\Users\Abhinav\Desktop\translated.txt")


wb = xlrd.open_workbook(loc)
sh = wb.sheet_by_index(0)

fin = open(fin_loc, "rt")
fout=open(fout_loc,"wt")

for i in range(1000): #converting excel sheet to dictionary
    cell_value_class = sh.cell(i,0).value
    cell_value_id = sh.cell(i,1).value
    d[cell_value_class] = cell_value_id

def convert(s):
    flag=False
    starting=''
    ending=''
    temp=''
    for i in s:
        if i.isalpha():
            flag=True
            temp+=i
        else:
            if flag==False:
                starting+=i
            elif flag:
                ending+=i
    if temp.lower() in d.keys():
        temp=d[temp.lower()]
        temp=starting+temp+ending
        return temp
    else:
        return s


for line in fin:

    for word in line.split():
        if word.isalpha():
            if word.lower() in d.keys():
                line=line.replace(word,d[word.lower()])
        else:
            temp=convert(word)
            line = line.replace(word, temp)


    fout.write(line)


