import xlrd
import xlwt
loc=("suru1.xlsx")
wa=xlwt.Workbook()
wb=xlrd.open_workbook(loc)
ws1=wa.add_sheet("Data")
ws=wb.sheet_by_index(0)
d={}
l1=[]
for i in range(1,ws.nrows):
  a=ws.cell_value(i,0)
  l1.append(a)
d["name"]=l1
l=[]
for i in range(1,ws.nrows):
  a=ws.cell_value(i,1)
  l.append(a)
d["pass"]=l
l2=[]
for i in range(1,ws.nrows):
  a=ws.cell_value(i,2)
  l2.append(a)
d["bal"]=l2
print(d)
task="yes"
while(task=="yes"):
  print("enter your name")
  name=input()
  for i in range(0,len(l1)):
    if(name==l[i]):
      print("enter your pass")
      password=input()
      for j in range(0,len(l)):
        if(password==l[i]):
          print("1. Debit")
          print("2. Credit")
          print("3. Money transfer")
          print("4. Check Balance")
          print("5. Password Change")
          n=int(input())
          if(n==1):
            print("enter ammount to Debit")
            c=int(input())
            if(c<=l2[i]):
              print("successful")
              d["bal"][i]=ws.cell_value(i+1,2)-c
              print(d["bal"][i])
              #ws1.write(i+1,2,d["bal"][i])
            else:
              print("not sufficient balance")
          elif(n==2):
            print("enter ammount to credit")
            c=int(input())
            print("successful")
            d["bal"][i]=ws.cell_value(i+1,2)+c
            print(d["bal"][i])
            #ws1.write(i+1,2,d["bal"][i])
          elif(n==3):
            print("enter name to transfer money")
            x=input()
            for k in range(0,len(l1)):
              if(x==l1[k]):
                tm=int(input("enter the money\n"))
                d["bal"][k]=ws.cell_value(k+1,2)+tm
                d["bal"][i]=ws.cell_value(i+1,2)-tm
                print(d["bal"][i])
                #ws1.write(i+1,2,d["bal"][i])
                print(d["bal"][k])
                #ws1.write(k+1,2,d["bal"][k])
                break
            
          elif(n==4):
            print("your balance is now:\n",d["bal"][i])
          elif(n==5):
            ps=input("enter current password\n")
            if(ps==l[i]):
              d["pass"][i]=input("enter new password\n")
              #ws1.write(i+1,2,d["pass"][i])
            else:
              print("password is worng")
          else:
            print("please enter right condition")
          break
        else:
          print("worng password")
          break
      else:
       print("user not found")
       break
  #wa.save("project.xlsx")
  print(d)
  task=input("do you want open again:- yes or no\n")
ws1.write(0,0,"name")
ws1.write(0,1,"pass")
ws1.write(0,2,"bal")
for i in range(0,len(d["name"])):
  ws1.write(i+1,0,d["name"][i])
  ws1.write(i+1,1,d["pass"][i])
  ws1.write(i+1,2,d["bal"][i])
wa.save("suru1.xlsx")