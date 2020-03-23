import os
import shutil
from shutil import copyfile
from pathlib import Path
import datetime
from color_class import bcolors

year = datetime.date.today().year
#year = 2019

def move_oe(projectCode,projectNumber,accountNumber):
    #move old TAB files to OUT Folder
    out_src = os.path.join(r'P:\TAB\Tab Export',str(year),projectNumber,accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'OE_C.xlsx')
    out_dst = os.path.join(r'P:\TAB\Tab Export',str(year),projectNumber,'OUT',accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'OE_C.xlsx')
    
    #move from export to tab
    src = os.path.join(r'P:\ONLINE DATA\Online Data Export',accountNumber,projectCode,accountNumber+'_'+projectCode+'_'+'Open-End_C.xlsx')
    tab_dst = os.path.join(r'P:\TAB\Tab Export',str(year),projectNumber,accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'OE_C.xlsx')
    
    print('.')
    print('.')
    print('.')
    print(bcolors.UNDERLINE+'Move Old OE file from TAB to OUT'+bcolors.ENDC)
    print('OUT Source path --------------->',out_src)
    print('OUT Destination path ------>',out_dst)
    print('-----------------------------------------------------')

    try:
        copyfile(out_src,out_dst)
        print(bcolors.OKBLUE+'Files have been moved succesfully'+bcolors.ENDC)
        
    except Exception as e:
        print(bcolors.FAIL+'Something went wrong when moving OE to OUT folder.'+bcolors.ENDC)
        print(bcolors.FAIL+str(e)+bcolors.ENDC)

    #move OE to TAB folder
    print('.')
    print('.')
    print('.')
    print(bcolors.UNDERLINE+'Move Open End file to TAB folder'+bcolors.ENDC)
    print('Source path --------------->',src)
    print('Destination TAB path ------>',tab_dst)
    print('-----------------------------------------------------')

    #perform the copy
    try:
        copyfile(src,tab_dst)
        print(bcolors.OKBLUE+'Files have been moved succesfully'+bcolors.ENDC)
        return True
    except Exception as e:
        print(bcolors.FAIL+'File move for OE unsuccessful.'+bcolors.ENDC)
        print(bcolors.FAIL+str(e)+bcolors.ENDC)
        return False

def move_carded(projectCode,projectNumber,accountNumber):
    #move from export to wincross
    src = os.path.join(r'P:\ONLINE DATA\Online Data Export',accountNumber,projectCode,accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'CEBIN_C.TXT')
    wc_dst = os.path.join(r'P:\ONLINE DATA\Online Docs\Online Wincross Data',str(year),projectNumber,accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'CEBIN_C.TXT')

    #move old files from TAB to OUT
    out_src = os.path.join(r'P:\TAB\Tab Export',str(year),projectNumber,accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'CEBIN_C.TXT')
    out_dst = os.path.join(r'P:\TAB\Tab Export',str(year),projectNumber,'OUT',accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'CEBIN_C.TXT')

    #move old carded into out
    print('-------------------------------')
    print(bcolors.UNDERLINE+'Move Carded from TAB to OUT'+bcolors.ENDC)
    print('Source path --------------->',out_src)
    print('Destination Wincross path ->',out_dst)
    try:
        copyfile(out_src,out_dst)
        print(bcolors.OKBLUE+'Files have been moved succesfully'+bcolors.ENDC)
    except Exception as e:
        print(bcolors.FAIL+'File move unsuccessful for Carded TAB to OUT.'+bcolors.ENDC)
        print(bcolors.FAIL+str(e)+bcolors.ENDC)
       
    print('-------------------------------')

    #print paths
    print('-------------------------------')
    print(bcolors.UNDERLINE+'Move Carded to Wincross Folder'+bcolors.ENDC)
    print('Source path --------------->',src)
    print('Destination Wincross path ->',wc_dst)

    #perform the copy of export to Wincross
    try:
        copyfile(src,wc_dst)
        print(bcolors.OKBLUE+'Files have been moved succesfully'+bcolors.ENDC)
        
    except Exception as e:
        print(bcolors.FAIL+'File move for Carded to Wincross unsuccessful.'+bcolors.ENDC)
        print(bcolors.FAIL+str(e)+bcolors.ENDC)
        
    print('-------------------------------')

    #move to tab
    src = os.path.join(r'P:\ONLINE DATA\Online Data Export',accountNumber,projectCode,accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'CEBIN_C.TXT')
    tab_dst = os.path.join(r'P:\TAB\Tab Export',str(year),projectNumber,accountNumber+'_'+projectCode+'_'+projectNumber+'_'+'CEBIN_C.TXT')

    print('.')
    print('.')
    print('.')
    print('-------------------------------')
    print(bcolors.UNDERLINE+'Move Carded to TAB Folder'+bcolors.ENDC)
    print('Source path --------------->',src)
    print('Destination TAB path ------>',tab_dst)

    #perform copy of carded to TAB
    try:
        copyfile(src,tab_dst)
        print(bcolors.OKBLUE+'Files have been moved from export to TAB succesfully'+bcolors.ENDC)
        return True
    except Exception as e:
        print(bcolors.FAIL+'File move for Carded to TAB unsuccessful.'+bcolors.ENDC)
        print(bcolors.FAIL+str(e)+bcolors.ENDC)
        return False
    print('-------------------------------')