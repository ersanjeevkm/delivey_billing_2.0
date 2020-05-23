#Billing 2.0 python2

import os
from os import system
from time import sleep
import getpass
import csv
from openpyxl import load_workbook
import openpyxl
import operator

#nos to string
def Words(n):
 units = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
 teens = ["Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen","Seventeen","Eighteen","Nineteen"]
 tens = ["Twenty","Thirty", "Fourty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
 if n <=9:
   return units[n]
 elif n >= 10 and n <= 19:
   return teens[n-10]
 elif n >= 20 and n <= 99:
   return tens[(n//10)-2] + " " + (units[n % 10] if n % 10 !=0 else "")
 elif n >= 100 and n <= 999:
   return Words(n//100) + " Hundred " + (Words(n % 100) if n % 100 !=0 else "")
 elif n >= 1000 and n <= 99999:
   return Words(n//1000) + " Thousand " + (Words(n % 1000) if n % 1000 !=0 else "")
 elif n >= 100000 and n <= 9999999:
   return Words(n//100000) + " Lakh " + (Words(n % 100000) if n % 100000 !=0 else "")
 elif n >= 10000000:
   return Words(n//10000000) + " Crore " + (Words(n % 10000000) if n % 10000000 !=0 else "")



i = "--------------------------------------------"
edit = False
in_des = ''
while(True):
  #login
  print("Welcome to Sanjeev delivery chalan system\n"+i)
  login = open("D:/delivery billing/login.txt")
  users = login.readline().strip()
  paswrd = login.readline().strip()
  crt_usr = ""
  while(True):
    a = raw_input("Enter username : ")
    if a.lower().strip()==users:
      crt_usr = a.lower().strip()
     
      b = getpass.getpass()
      if b==paswrd:
        print("Login succeded\nWelcome "+a)
        break
      else:
        print("Password incorrect ! Try again\n");
    else:
      print("Username incorrect ! Try again\n");

  no = login.readline().strip()
  print("\nOld bill no. : "+no+"\n"+i)
  no = int(no)
  sleep(2)
  system("cls")
  login.close()
  #menus
  while(True):
    print("Welcome to billing menu : \n")
    print("\nOld bill no. : "+str(no)+"\n"+i+i+i)
    print("-> 1 - New bill\n-> 2 - Add party\n-> 3 - Add product\n\n"+i+i+i+"\n-> 4 - Edit bill\n-> 5 - Edit party\n-> 6 - Edit products\n\n-> 7 - Print previous DC's\n\n"+i+i+i+"\n-> 8 - Switch to other DC\n")
    option = raw_input("Pick the desired option (\'n\' to exit) : ").strip()
    if option.lower() == "n":
      break
    if(option<"1")or(option>"8"):
      print(i+"\nPick the correct option (\'n\' to exit)\n")
      continue
    sleep(1)
    system("cls")

    #new bill
    if(option=='1'):
      print("---------------------------- NEW BILL --------------------------\n")
      print("----------------------- CHOOSE PARTY ---------------------------\n")
      print(i+i+i)
      crt_party=[]
      crt_prods=[]

      with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_party.csv") as party_files:
        parties = csv.reader(party_files)
        next(parties)
        for j in parties:
          print "  {:2} -> {:30}".format(j[0],j[1])," | ",
          try:
            j = next(parties)
          except:
            print("")
            continue
          print "{:2} -> {:30}".format(j[0],j[1])," | ",
          try:
            j = next(parties)
          except:
            print("")
            continue
          print("{:2} -> {:30}".format(j[0],j[1]))
          
          print(i+i+i)
      
      ch_party = raw_input("\nChoose party : ")
      with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_party.csv") as party_files:
        parties = csv.reader(party_files)
        next(parties)  
        for j in parties:
          if(ch_party==j[0]):
            crt_party.append(j)
            break

      print("\nCurrent party: ")
      print(" -> "+crt_party[0][1])
      print(i+i+i)

      sleep(1)
      system("cls")

      print("----------------------- CHOOSE PRODUCTS ---------------------------\n")
      print(i+i+i)
      with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_prod.csv") as prod_files:
        prod = csv.reader(prod_files)
        next(prod)
        for j in prod:
          print "  {:2} -> {:30}".format(j[0],j[1])," | ",
          try:
            j = next(prod)
          except:
            print("")
            continue
          print "{:2} -> {:30}".format(j[0],j[1])," | ",
          try:
            j = next(prod)
          except:
            print("")
            continue
          print "{:2} -> {:30}".format(j[0],j[1])

          print(i+i+i)

      while(True):
        ch_prod = raw_input("\nChoose product : ")
        quant = raw_input("Enter quantity : ")
        
        with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_prod.csv") as prod_files:
          prods = csv.reader(prod_files)
          next(prods)  
          for j in prods:
            if(ch_prod==j[0]):
              j.append(quant)
              print("Price/"+j[3]+" = "+j[4])
              crt_prods.append(j)
              break
        ano = raw_input("Choose another product : (y/n) ")
        if(ano.lower() == "y"):
          continue
        else:
          break
      total_quant = 0
      total_val = 0
      print("Choosed products : ")
      for j in crt_prods:

        total_quant += float(j[5])
        total_val += float(j[5])*float(j[4])
        print(" -> "+j[1])
      vehi_no = raw_input("\nEnter vehicle no. : ")
      no+=1
      #writing the excel sheet
      workbook = load_workbook(filename="D:/delivery billing/"+crt_usr+"/template.xlsx")
      sheet = workbook.active
      sheet["C10"] = no
      sheet["F12"] = crt_party[0][4]
      sheet["A19"] = crt_party[0][1]
      sheet["A20"] = crt_party[0][2]
      sheet["A21"] = crt_party[0][3]
      sheet["A22"] = crt_party[0][4]
      sheet["B23"] = crt_party[0][5]
      sheet["B24"] = crt_party[0][6]
      sheet["F9"] = crt_party[0][7]
      count = 0
      for k in crt_prods:
        count+=1
        sheet["B{}".format(26+count)] = count
        sheet["C{}".format(26+count)] = k[1]
        sheet["D{}".format(26+count)] = k[2]
        sheet["F{}".format(26+count)] = k[3]
        sheet["E{}".format(26+count)] = k[5]
        sheet["G{}".format(26+count)] = float(k[4])*float(k[5])
      sheet["F10"] = vehi_no
      sheet["E38"] = round(total_quant,3)
      sheet["G38"] = int(round(total_val))
      sheet["B40"] = Words(int(round(total_val)))+" Only."
      
      #update bill no in txt file
      if edit == False:
        login = open("D:/delivery billing/login.txt","r")
        ch_no = login.readlines()
        ch_no[2] = str(no)
      
        login = open("D:/delivery billing/login.txt","w")
        login.writelines(ch_no)
        login.close()
      workbook.save(filename="C:/Users/<PC usr name>/Desktop/delivery bills/{}.{}.xlsx".format(no,crt_party[0][1]))
      
      #print bill
      oprint = raw_input("\nPrint the bill ? (y/n) ")
      if(oprint.lower() == 'y'):
        os.startfile("C:/Users/<PC usr name>/Desktop/delivery bills/{}.{}.xlsx".format(no,crt_party[0][1]),'print')
        
    #add party
    if(option=='2'):
      print("------------------- ADD PARTY -------------------\n"+i+i+i)
      while(True):
        new_part={} 
        new_part["Name"] = raw_input("-> Enter party name : ")
        new_part["Add_line1"] = raw_input("\n-> Enter add line 1 : ")
        new_part["Add_line2"] = raw_input("\n-> Enter add line 2 : ")
        new_part["Add_line3"] = raw_input("\n-> Enter add line 3 : ")
        new_part["GSTIN"] = raw_input("\n-> Enter GSTIN : ")
        new_part["CODE"] = raw_input("\n-> Enter CODE no. : ")
        new_part["Purpose"] = raw_input("\n-> Purpose of transportation : ")

        sno=0
        with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_party.csv") as party_file:
          for row in party_file:
            sno+=1;
        print("\nTotal parties : "+str(sno))    
        keys=["S.No","Name","Add_line1","Add_line2","Add_line3","GSTIN","CODE","Purpose"]
        with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_party.csv",'a+') as party_file:
          new_part["S.No"] = sno
          write_party = csv.DictWriter(party_file,fieldnames=keys,lineterminator='\n') 
          write_party.writerow(new_part)

        party_file = open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_party.csv")
        party_data = csv.reader(party_file)
        next(party_data)
        sorted_party = sorted(party_data, key=operator.itemgetter(1), reverse=False)

        party_file = open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_party.csv","w")
        party_write = csv.writer(party_file,lineterminator='\n')
        sno = 0
        party_write.writerow(keys)
        for l in sorted_party:
          sno+=1
          l[0] = sno
          party_write.writerow(l)
        party_file.close()
        
        cont = raw_input("\nAdd another party? (y/n) ")
        if(cont.lower()=='n'):
          break
        elif(cont.lower()=='y'):
          sleep(1)
          system("cls")

    #add product
    if(option=='3'):
      print("-------------- ADD PRODUCT --------------\n"+i+i+i)
      while(True):
        new_prod={}
        new_prod["Name"] = raw_input("-> Enter product name : ")
        new_prod["SCN"] = raw_input("\n-> Enter SCN code : ")
        new_prod["UOM"] = raw_input("\n-> Enter UOM : ")
        new_prod["Price"] = raw_input("\n-> Enter Price/"+new_prod["UOM"]+" : ")
        
        sno=0
        with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_prod.csv") as prod_file:
          for row in prod_file:
            sno+=1;
        print("\nTotal Products : "+str(sno))   
        
        keys=["S.No","Name","SCN","UOM","Price"]
        with open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_prod.csv",'a+') as prod_file:
          new_prod["S.No"] = sno
          write_prod = csv.DictWriter(prod_file,fieldnames=keys,lineterminator='\n')
          write_prod.writerow(new_prod)
          
        party_file = open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_prod.csv")
        party_data = csv.reader(party_file)
        next(party_data)
        sorted_party = sorted(party_data, key=operator.itemgetter(1), reverse=False)

        party_file = open("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_prod.csv","w")
        party_write = csv.writer(party_file,lineterminator='\n')
        sno = 0
        party_write.writerow(keys)
        for l in sorted_party:
          sno+=1
          l[0] = sno
          party_write.writerow(l)
        party_file.close()
          
        cont = raw_input("Add another product? (y/n) ")
        if(cont.lower()=='n'):
          break
        elif(cont.lower()=='y'):
          sleep(1)
          system("cls")

    if(option=="4"):
      bills=[]
      print("\n-------------- EDIT BILL --------------\n\n"+i+i+i)
      for bill in sorted(os.listdir("C:/Users/<PC usr name>/Desktop/delivery bills/")):
        bills.append(bill.split(".xlsx")[0])
      newline = 0
      for name in bills:
        newline+=1
        if newline%3 == 0:
          print(" {:34}".format(name))
          print(i+i+i)
          continue
        print " {:34}".format(name)," |",
        
      while(True):
        billno = raw_input("\n\nEnter bill no. ")
        for name in bills:
          if(name.split(".")[0]) == billno:
            os.startfile("C:/Users/<PC usr name>/Desktop/delivery bills/"+name+'.xlsx')
            break
        bquest = raw_input("Edit another bill ? (y/n) ")
        if(bquest.lower() == 'n'):
          break

    if(option=="5"):
      os.startfile("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_party.csv")

    if(option=="6"):
      os.startfile("D:/delivery billing/"+crt_usr+"/"+crt_usr+"_prod.csv")

    if(option=="7"):
      bills=[]
      print("\n-------------- PRINT DC's --------------\n\n"+i+i+i)
      for bill in sorted(os.listdir("C:/Users/<PC usr name>/Desktop/delivery bills/")):
        bills.append(bill.split(".xlsx")[0])
      newline = 0
      for name in bills:
        newline+=1
        if newline%3 == 0:
          print(" {:34}".format(name))
          print(i+i+i)
          continue
        print " {:34}".format(name)," |",

      while(True):
        billno = raw_input("\n\nEnter DC no. ")
        for name in bills:
          if(name.split(".")[0]) == billno:
            os.startfile("C:/Users/<PC usr name>/Desktop/delivery bills/"+name+'.xlsx','print')
            break
          
        bquest = raw_input("Print another DC ? (y/n) ")
        if(bquest.lower() == 'n'):
          break


    print(i+i+i)
    
    if(option=="8"):
      print("\n-------------- SWITCH TO DC No. --------------\n\n"+i+i+i)
      no1 = raw_input("Enter DC no. : ")
      no = int(no1)-1
      sleep(1)
      system("cls")
      edit = True
      continue

    #final cond
    des = raw_input("\nBack to bill menu ? (y/n) ")
    if(des.lower() == 'n'):
      break
    else:
      if edit == True:
        edit = False
        login = open("D:/delivery billing/login.txt")
        login.readline()
        login.readline()
        no = login.readline().strip()
        no = int(no)
        login.close()
        
      sleep(1)
      system("cls")
      continue
  break
