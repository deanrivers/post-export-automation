import os
import datetime
from color_class import bcolors

year = datetime.date.today().year

def create_folders(project_number):
    try:
        #Wincross variables
        wc_dir = os.path.join(r'P:\ONLINE DATA\Online Docs\Online Wincross Data',str(year),project_number)
        # wc_folder_exists = os.path.exists(wc_dir)
        os.mkdir(wc_dir)
        print(bcolors.OKBLUE+'Wincross folder created for:'+bcolors.ENDC,project_number)
    except Exception as e:
        print('.')
        print('.')
        print('.')
        print('The Wincross folder already exists')
        print(bcolors.RED+str(e)+bcolors.ENDC)

    try:
        #TAB variables
        tab_dir = os.path.join(r'P:\TAB\Tab Export',str(year),project_number)
        tab_out_dir = os.path.join(r'P:\TAB\Tab Export',str(year),project_number,'OUT')
        # tab_folder_exists = os.path.exists(tab_dir)
        os.mkdir(tab_dir)
        os.mkdir(tab_out_dir)
        print(bcolors.OKBLUE+'TAB folder created for:'+bcolors.ENDC,project_number)

    except Exception as e:
        print('.')
        print('.')
        print('.')
        print('This TAB folder already exists.')
        print(bcolors.RED+str(e)+bcolors.ENDC)