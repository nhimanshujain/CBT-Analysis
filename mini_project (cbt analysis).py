#mini project
'''CBT analysis(Reading data from excel files and generating class average, finding highest,
lowest and generating reports) '''

import xlrd

#open workbook
workbook=xlrd.open_workbook('mini_project.xlsx')

#open worksheet
worksheet=workbook.sheet_by_index(0)

#total number of rows and columns
total_rows=worksheet.nrows
total_cols=worksheet.ncols

print(' ',format(' ','<50'),'This program will generate a class report\n')

#title of xlsx sheet
print(' ',format(' ','<55'),worksheet.cell_value(0,0),'\n')


#bio
def bio():
        print('\n\n',format(' ','<68'),'Engineering Biology')
        bt_marks=[worksheet.cell_value(i,3) for i in range(2,total_rows)]
        print('\nClass average=',int(sum(bt_marks)/len(bt_marks)))
        
        print('\nHighest marks=',int(max(bt_marks)))
        highest_marks=[worksheet.cell_value(i+2,2) for i in range(len(bt_marks)) if bt_marks[i]==max(bt_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(bt_marks)) if bt_marks[i]==max(bt_marks)]
        print('Students securing',int(max(bt_marks)),':')
        j=0
        for i in highest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1
        
        print('\nLowest marks=',int(min(bt_marks)))
        lowest_marks=[worksheet.cell_value(i+2,2) for i in range(len(bt_marks)) if bt_marks[i]==min(bt_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(bt_marks)) if bt_marks[i]==min(bt_marks)]
        print('Students securing',int(min(bt_marks)),':')
        j=0
        for i in lowest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1

        def list_bt():
                data={}
                for i in bt_marks:
                        num=int(i)
                        if num not in data:
                                data[num]=0
                        data[num]+=1
                for i in sorted(data):
                        print('Number of students securing',i,'marks =',data[i])

        a=input('''\nIf you want complete detail of the class
                        press ==>y
                        else
                        press ==>n\n''')
        if a=='y':
                list_bt()
        

#python
def py():
        print('\n\n',format(' ','<55'),'Introduction To Computer Science Using Python')
        py_marks=[worksheet.cell_value(i,4) for i in range(2,total_rows)]
        print('\nClass average=',int(sum(py_marks)/len(py_marks)))
        
        print('\nHighest marks=',int(max(py_marks)))
        highest_marks=[worksheet.cell_value(i+2,2) for i in range(len(py_marks)) if py_marks[i]==max(py_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(py_marks)) if py_marks[i]==max(py_marks)]
        print('Students securing',int(max(py_marks)),':')
        j=0
        for i in highest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1
        
        print('\nLowest marks=',int(min(py_marks)))
        lowest_marks=[worksheet.cell_value(i+2,2) for i in range(len(py_marks)) if py_marks[i]==min(py_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(py_marks)) if py_marks[i]==min(py_marks)]
        print('Students securing',int(min(py_marks)),':')
        j=0
        for i in lowest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1

        def list_py():
                data={}
                for i in py_marks:
                        num=int(i)
                        if num not in data:
                                data[num]=0
                        data[num]+=1
                for i in sorted(data):
                        print('Number of students securing',i,'marks =',data[i])

        a=input('''\nIf you want complete detail of the class
                        press ==>y
                        else
                        press ==>n\n''')
        if a=='y':
                list_py()
        

#mechanics
def mec():
        print('\n\n',format(' ','<68'),'Engineering Mechanics')
        mec_marks=[worksheet.cell_value(i,5) for i in range(2,total_rows)]
        print('\nClass average=',int(sum(mec_marks)/len(mec_marks)))
        
        print('\nHighest marks=',int(max(mec_marks)))
        highest_marks=[worksheet.cell_value(i+2,2) for i in range(len(mec_marks)) if mec_marks[i]==max(mec_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(mec_marks)) if mec_marks[i]==max(mec_marks)]
        print('Students securing',int(max(mec_marks)),':')
        j=0
        for i in highest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1
        
        print('\nLowest marks=',int(min(mec_marks)))
        lowest_marks=[worksheet.cell_value(i+2,2) for i in range(len(mec_marks)) if mec_marks[i]==min(mec_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(mec_marks)) if mec_marks[i]==min(mec_marks)]
        print('Students securing',int(min(mec_marks)),':')
        j=0
        for i in lowest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1

        def list_mec():
                data={}
                for i in mec_marks:
                        num=int(i)
                        if num not in data:
                                data[num]=0
                        data[num]+=1
                for i in sorted(data):
                        print('Number of students securing',i,'marks =',data[i])

        a=input('''\nIf you want complete detail of the class
                        press ==>y
                        else
                        press ==>n\n''')
        if a=='y':
                list_mec()



#chemistry
def chm():
        print('\n\n',format(' ','<68'),'Engineering Chemistry')
        chm_marks=[worksheet.cell_value(i,6) for i in range(2,total_rows)]
        print('\nClass average=',int(sum(chm_marks)/len(chm_marks)))
        
        print('\nHighest marks=',int(max(chm_marks)))
        highest_marks=[worksheet.cell_value(i+2,2) for i in range(len(chm_marks)) if chm_marks[i]==max(chm_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(chm_marks)) if chm_marks[i]==max(chm_marks)]
        print('Students securing',int(max(chm_marks)),':')
        j=0
        for i in highest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1
        
        print('\nLowest marks=',int(min(chm_marks)))
        lowest_marks=[worksheet.cell_value(i+2,2) for i in range(len(chm_marks)) if chm_marks[i]==min(chm_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(chm_marks)) if chm_marks[i]==min(chm_marks)]
        print('Students securing',int(min(chm_marks)),':')
        j=0
        for i in lowest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1

        def list_chm():
                data={}
                for i in chm_marks:
                        num=int(i)
                        if num not in data:
                                data[num]=0
                        data[num]+=1
                for i in sorted(data):
                        print('Number of students securing',i,'marks =',data[i])

        a=input('''\nIf you want complete detail of the class
                        press ==>y
                        else
                        press ==>n\n''')
        if a=='y':
                list_chm()
        

#electronics
def elc():
        print('\n\n',format(' ','<65'),'Basic Electronics Engineering')
        elc_marks=[worksheet.cell_value(i,7) for i in range(2,total_rows)]
        print('\nClass average=',int(sum(elc_marks)/len(elc_marks)))
        
        print('\nHighest marks=',int(max(elc_marks)))
        highest_marks=[worksheet.cell_value(i+2,2) for i in range(len(elc_marks)) if elc_marks[i]==max(elc_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(elc_marks)) if elc_marks[i]==max(elc_marks)]
        print('Students securing',int(max(elc_marks)),':')
        j=0
        for i in highest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1
        
        print('\nLowest marks=',int(min(elc_marks)))
        lowest_marks=[worksheet.cell_value(i+2,2) for i in range(len(elc_marks)) if elc_marks[i]==min(elc_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(elc_marks)) if elc_marks[i]==min(elc_marks)]
        print('Students securing',int(min(elc_marks)),':')
        j=0
        for i in lowest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1

        def list_elc():
                data={}
                for i in elc_marks:
                        num=int(i)
                        if num not in data:
                                data[num]=0
                        data[num]+=1
                for i in sorted(data):
                        print('Number of students securing',i,'marks =',data[i])

        a=input('''\nIf you want complete detail of the class
                        press ==>y
                        else
                        press ==>n\n''')
        if a=='y':
                list_elc()
        
       
#maths
def math():
        print('\n\n',format(' ','<68'),'Engineering Mathematics')
        math_marks=[worksheet.cell_value(i,8) for i in range(2,total_rows)]
        print('\nClass average=',int(sum(math_marks)/len(math_marks)))
        
        print('\nHighest marks=',int(max(math_marks)))
        highest_marks=[worksheet.cell_value(i+2,2) for i in range(len(math_marks)) if math_marks[i]==max(math_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(math_marks)) if math_marks[i]==max(math_marks)]
        print('Students securing',int(max(math_marks)),':')
        j=0
        for i in highest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1
        
        print('\nLowest marks=',int(min(math_marks)))
        lowest_marks=[worksheet.cell_value(i+2,2) for i in range(len(math_marks)) if math_marks[i]==min(math_marks)]
        snr=[worksheet.cell_value(i+2,1) for i in range(len(math_marks)) if math_marks[i]==min(math_marks)]
        print('Students securing',int(min(math_marks)),':')
        j=0
        for i in lowest_marks:
                print('\t','Name:',i,"==>",'SNR:',snr[j])
                j+=1

        def list_math():
                data={}
                for i in math_marks:
                        num=int(i)
                        if num not in data:
                                data[num]=0
                        data[num]+=1
                for i in sorted(data):
                        print('Number of students securing',i,'marks =',data[i])

        a=input('''\nIf you want complete detail of the class
                        press ==>y
                        else
                        press ==>n\n''')
        if a=='y':
                list_math()
        


def details():
        snr=input('Enter your SNR=')
        for i in range(2,total_rows):
                a=worksheet.cell_value(i,1)
                if snr==a:
                        print('\nSNR=',snr)
                        print('Name=',worksheet.cell_value(i,2))
                        print('Biology marks=',int(worksheet.cell_value(i,3)))
                        print('Python marks=',int(worksheet.cell_value(i,4)))
                        print('Mechanics marks=',int(worksheet.cell_value(i,5)))
                        print('Chemistry marks=',int(worksheet.cell_value(i,6)))
                        print('Electronics marks=',int(worksheet.cell_value(i,7)))
                        print('Maths marks=',int(worksheet.cell_value(i,8)))

user=input('''To get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')

boo=False
while not boo:
        if user=='b':
                bio()
                user=input('''\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')
                if user=='b':
                        boo=False
                elif user=='q':
                        boo=True        
        elif user=='p':
                py()
                user=input('''\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')
                if user=='p':
                        boo=False
                elif user=='q':
                        boo=True
        elif user=='m':
                mec()
                user=input('''\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')
                if user=='m':
                        boo=False
                elif user=='q':
                        boo=True      
        elif user=='c':
                chm()
                user=input('''\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')
                if user=='c':
                        boo=False
                elif user=='q':
                        boo=True
        elif user=='e':
                elc()
                user=input('''\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')
                if user=='e':
                        boo=False
                elif user=='q':
                        boo=True      
        elif user=='ma':
                math()
                user=input('''\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')
                if user=='ma':
                        boo=False
                elif user=='q':
                        boo=True      
        elif user=='s':
                details()
                user=input('''\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')
                if user=='s':
                        boo=False
                elif user=='q':
                        boo=True      
        elif user=='q':
                boo=True
        else:
                print('\nInvalid input!')
                user=input('''\n\nTo get the details of :\nBiology ==> press b
Python ==> press p
Mechanics ==> press m
Chemistry ==> press c
Electronics ==> press e
Maths ==> press ma
Student details ==> press s
To quit ==> press q\n''')

                




