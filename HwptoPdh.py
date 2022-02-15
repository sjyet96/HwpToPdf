from tkinter import *
from tkinter import filedialog
import win32com.client as win32
import pandas as pd
from datetime import date, datetime

def pathSelctBt():
    global excelpath
    excelpath = filedialog.askopenfilename()
    expathent.delete(0,END)
    expathent.insert(0,excelpath)
    print(excelpath)
    
    
def hwppathSelctBt():
    global hwppath
    hwppath = filedialog.askopenfilename()
    hwppathent.delete(0,END)
    hwppathent.insert(0,hwppath)
    print(hwppath)
        
    
def toPDF():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = True

    hwp.Open(hwppath,"HWP","template:TRUE")

    field = hwp.GetFieldList(0,None).split("\x02")

    print(field)

    df=pd.read_excel(excelpath)


    name = "송지유"
    level = '선임'
    path = filedialog.askdirectory()
    middlename='접수증_제출자보관용_'
    file = path + middlename
    today = date.today()
    usertime = datetime.today()
    pathlist=[]
    for i in range(len(df)):
        hwp.PutFieldText('접수번호', df['접수번호'][i])
        #hwp.PutFieldText('접수일시', str(today.year)+'.'+str(today.month)+'.'+str(today.day)+'.'+" "+str(usertime.hour)+':'+str(usertime.minute)+':'+str(usertime.second))
        hwp.PutFieldText('과제명', df['과제명'][i])
        hwp.PutFieldText('도입기업명', df['도입기업명'][i])
        hwp.PutFieldText('총괄책임자', df['도입기업 책임자'][i])
        hwp.PutFieldText('공급기업명', df['공급기업명'][i])
        hwp.PutFieldText('책임자', df['책임자'][i])
        hwp.PutFieldText('총사업비', format(df['총사업비'][i],','))
        hwp.PutFieldText('정부지원금', format(df['정부지원금'][i],','))
        hwp.PutFieldText('민간부담금', format(df['민간부담금'][i],','))
        hwp.PutFieldText('직급', levelent.get())
        hwp.PutFieldText('이름', manegerent.get())
        hwp.PutFieldText('날짜', today.day)
        hwp.SaveAs(path+"/"+ "접수증_"+str(df['도입기업명'][i])+".pdf" , "PDF")
        pathlist.append(path+"/"+ "접수증_"+str(df['도입기업명'][i])+".pdf")
    
    expathent.delete(0,END)
    hwppathent.delete(0,END)
    manegerent.delete(0,END)
    levelent.delete(0,END)
    df['접수증 경로']=pathlist
    print(df)
    df.to_excel('D:/접수증/접수증리스트.xlsx')
 

root = Tk()
root.title("test")
#root. geometry("640x480")

expathbt = Button(root, text = "선택", command=pathSelctBt)
hwppathbt = Button(root, text = "선택", command=hwppathSelctBt)
pdfbt = Button(root, text = "pdf 변환", command=toPDF)

l0 = Label(root, text = "엑셀 파일")
l01 = Label(root, text = "한글양식")
l1 = Label(root, text = "담당자")
l2 = Label(root, text = "직급")


expathent = Entry(root, width =10)
hwppathent = Entry(root, width =10)
manegerent = Entry(root, width =10)
levelent = Entry(root, width =10)



expathent.grid(row = 0,column=1)
expathbt.grid(row = 0,column=2)
l0.grid(row = 0,column=0)

l01.grid(row = 1,column=0)
hwppathent.grid(row = 1,column=1)
hwppathbt.grid(row = 1,column=2)

l1.grid(row = 2,column=0)
manegerent.grid(row = 2,column=1)
l2.grid(row = 3,column=0)
levelent.grid(row = 3,column=1)
pdfbt.grid(row = 4,column=1)
#levelent.pack()


#txt1 = Text(root, width =30, height = 10)
#bt1 = Button(root, text = "CLICK", command=btn)
root.mainloop()
    
        



