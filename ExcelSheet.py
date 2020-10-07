#!/usr/bin/python

import os
import os.path     
import openpyxl
import time
import platform
import terminal_banner
import termcolor

if(platform.system() == 'Windows'):
   os.system('cls')
if(platform.system() == 'Linux'):
   os.system('clear')   



header = """


 ________                                __   ______   __                              __
/        |                              /  | /      \ /  |                            /  |
$$$$$$$$/  __    __   _______   ______  $$ |/$$$$$$  |$$ |____    ______    ______   _$$ |_
$$ |__    /  \  /  | /       | /      \ $$ |$$ \__$$/ $$      \  /      \  /      \ / $$   |
$$    |   $$  \/$$/ /$$$$$$$/ /$$$$$$  |$$ |$$      \ $$$$$$$  |/$$$$$$  |/$$$$$$  |$$$$$$/
$$$$$/     $$  $$<  $$ |      $$    $$ |$$ | $$$$$$  |$$ |  $$ |$$    $$ |$$    $$ |  $$ | __
$$ |_____  /$$$$  \ $$ \_____ $$$$$$$$/ $$ |/  \__$$ |$$ |  $$ |$$$$$$$$/ $$$$$$$$/   $$ |/  |
$$       |/$$/ $$  |$$       |$$       |$$ |$$    $$/ $$ |  $$ |$$       |$$       |  $$  $$/
$$$$$$$$/ $$/   $$/  $$$$$$$/  $$$$$$$/ $$/  $$$$$$/  $$/   $$/  $$$$$$$/  $$$$$$$/    $$$$/



"""

desc = "ExcelSheet is a spreadsheet program."
dev_info = """
Python3
Developed by: Kumar Atul Jaiswal (@whoiskumaratul)
Copyright: ©2020 Hacking Truth
"""

banner = terminal_banner.Banner(header)
print(termcolor.colored(banner.text,'cyan'), end="")
print(termcolor.colored(desc,'white', attrs=['bold']), end = "")
print(termcolor.colored(dev_info,'yellow'))


print('Enter a Name for column A')
a = input()
print('Enter a Name for column B')
b = input()
print('Enter a Name for column C')
c = input()
print('Enter a Name for column D')
d = input()

print('\n')


print('Enter a Value for: ' + b)
bb = input()

print('Enter a value for: ' + c)
cc = input()

print('Enter a value for: ' + d)
dd = input()

print('\n')
try:
   addonce = input("[?] Do you want to add more value? Y/n: ")
   addonce =  addonce.lower()
   if addonce == "y" or addonce == "":
      print('Please wait...')
      time.sleep(1)
      print('[I] Column\'s Value Type more...')
      print('Enter a value for: ' + b)
      bbMore = input()
      print('Enter a value for: ' + c)
      ccMore = input()
      print('Enter a value for: ' + d)
      ddMore = input()

except:
      print("")

#Open a new file and add a user's data
e  = open('barchartwithuser.xlsx', 'w') 
#print(e.write(a))
#print(e.write(b))
#print(e.write(c))
#print(e.write(d))



print('\n')
try:
   addonce = input("[?] Do you want to add 6 more value? Y/n: ")
   addonce =  addonce.lower()
   if addonce == "y" or addonce == "":
      for abc in range(1, 2):
          abc = ['bbMoreMore', 'ccMoreMore', 'ddMoreMore', 'bbMoreMoreMore', 'ccMoreMoreMore', 'ddMoreMoreMore', 'bbMMMM', 'ccMMMM', 'ddMMMM', 'ee', 'ff', 'gg', 'hh', 'ii', 'jj', 'kk', 'll', 'mm',]
          print('Please wait...')
          time.sleep(1)
          print('[I] Column\'s Value Type more...')
          print('Enter a value for: ' + b)
          bbMoreMore = input()
          print('Enter a value for: ' + c)
          ccMoreMore = input()
          print('Enter a value for: ' + d)
          ddMoreMore = input()
          
          if abc == "ddMoreMore":
             break
          
          print('\n')      
          print('Enter a value for: ' + b)
          bbMoreMoreMore = input()
          print('Enter a value for: ' + c)
          ccMoreMoreMore = input()
          print('Enter a value for: ' + d)
          ddMoreMoreMore = input()     


          if abc == "ddMoreMoreMore":
             break
             
          print('\n')    
          print('Enter a value for: ' + b)
          bbMMMM = input()
          print('Enter a value for: ' + c)
          ccMMMM = input()
          print('Enter a value for: ' + d)
          ddMMMM = input()  
            
          if abc == "ddMMMM":
             break  
             
          print('\n')  
          print('Enter a value for: ' + b)
          ee = input()
          print('Enter a value for: ' + c)
          ff = input()
          print('Enter a value for: ' + d)
          gg = input() 


          if abc == "gg":
             break  
             
          print('\n')  
          print('Enter a value for: ' + b)
          hh = input()
          print('Enter a value for: ' + c)
          ii = input()
          print('Enter a value for: ' + d)
          jj = input() 


          if abc == "jj":
             break  
             
          print('\n')  
          print('Enter a value for: ' + b)
          kk = input()
          print('Enter a value for: ' + c)
          ll = input()
          print('Enter a value for: ' + d)
          mm = input() 








except:
      print("")




wb = openpyxl.Workbook()
sheet = wb.active
#print(sheet)


for i in range(1, 11):
   sheet['A' + str(i)] =  i


sheet['A1'] = a
sheet['B1'] = b
sheet['C1'] = c
sheet['D1'] = d

sheet['B2'] = bb
sheet['C2'] = cc
sheet['D2'] = dd

#sheet['A2'] = 1
#For add a row's and colum's value


try:
 
   sheet['B3'] = bbMore
   sheet['C3'] = ccMore
   sheet['D3'] = ddMore
   
   sheet['B4'] = bbMoreMore
   sheet['C4'] = ccMoreMore
   sheet['D4'] = ddMoreMore
   
   sheet['B5'] = bbMoreMoreMore
   sheet['C5'] = ccMoreMoreMore
   sheet['D5'] = ddMoreMoreMore
   
   sheet['B6'] = bbMMMM
   sheet['C6'] = ccMMMM
   sheet['D6'] = ddMMMM

   sheet['B7'] = ee
   sheet['C7'] = ff
   sheet['D7'] = gg

   sheet['B8'] = hh
   sheet['C8'] = ii
   sheet['D8'] = jj

   sheet['B9'] = kk
   sheet['C9'] = ll
   sheet['D9'] = mm


 
except:
    print("")


"""
try:
 
   sheet['B4'] = bbMoreMore
   sheet['C4'] = ccMoreMore
   sheet['D4'] = ddMoreMore
   
   sheet['B5'] = bbMoreMoreMore
   sheet['C5'] = ccMoreMoreMore
   sheet['D5'] = ddMoreMoreMore
   
   #sheet['B6'] = bbMoreMoreMore
   #sheet['C6'] = ccMoreMoreMore
   #sheet['D6'] = ddMoreMoreMore
   
except:
    print("")

"""



sheet.row_dimensions[1].height = 70
sheet.column_dimensions['B'].width = 30
sheet.column_dimensions['C'].width = 30
sheet.column_dimensions['D'].width = 30

print('Please wait..')
time.sleep(1)
print('File Successfully Created')

wb.save('check.xlsx')

