###################################################################################################
#Script Name      :main.py
#Description      :See below
#Author           :Dean Rivers
#Date Created     :1/28/20
#Last Updated     :2/11/20
###################################################################################################
'''
PLEASE READ:
This script should only be used for simple one-part studies which have this file structure.
Implementing Multi-part studies is next on the to do list.
Example Wincross Folder -> P:\ONLINE DATA\Online Docs\Online Wincross Data\2020\660120'
Example TAB folder ------> P:\TAB\Tab Export\2020\660120'
'''
#########################################################
'''
pip installs needed:
1. pip install -r requirements.txt
'''
#########################################################
'''
This script will be used to automate
the process of moving files around
after export and editing the Open End File.
You wil still have to manually remove the apostrophes in the OE file in the TAB folder.
'''
#########################################################
'''
Required Setup:
1. Export the raw datamap. (A project number HAS to be associated with the project.)
2. Export carded and OE.
'''
#########################################################
'''
Process of what is happening:
1. Carded data gets moved from Data export to Corresponding Wincross and TAB Folders. A carded file that already exist in TAB will be moved to the OUT folder.
2. OE gets moved into TAB and is cleaned (deletes typical columns and PII columns). An OE that already exists in TAB will be moved to the OUT folder. 
3. You will still have to open up the OE and remove any fields that may be unique to your study. You will also have to manually replace the apostrophes.
4. An email will be prepped in Outlook. You will have to add whoever needs to be added depending on the project team.
'''
###################################################################################################
###################################################################################################
import os
import datetime
import move_files
from excel_oe_edits import excel_work
import pathlib
import send_email
from color_class import bcolors
import sys
from create_folders import create_folders

def main(data_export_path,projectNumber):

    split = data_export_path.split('\\')

    #check that path for data export is valid
    try:
        #User input -> What is the Account Number for this Client?
        accountNumber = split[3]

        #User input -> What is the Project Code?
        projectCode = split[4]

        #get datamap filename
        #dm_filename = r'000200_01212000024_660120_DM.TXT'
        for filename in os.listdir(data_export_path):
            if('_DM.TXT' in filename):
                print(filename)
                dm_filename = filename

                #parse out the needed text from the DM name
                #start -> 'projectcode_'
                #end -> '-DM
                s = dm_filename
                start = s.find(str(projectCode)+'_') + len(str(projectCode)+'_')
                end = s.find('_DM')
                substring = s[start:end]

                print('Project Number:',projectNumber)
                print('Substring:',substring)

                if substring == projectNumber:
                    substringMatch = True
                else:
                    substringMatch = False
                break
        
        if substringMatch:
            projectNumberValid = True
        else:
            projectNumberValid = False
            return False, 'The project number must match the Datamap.'
        #projectNumberValid = False

    except Exception as e:
        print('-------------------------------')
        print(bcolors.FAIL + 'ERROR' + bcolors.ENDC)
        print(bcolors.RED+'The Data Export Path provided is invalid.'+bcolors.ENDC)
        print(bcolors.RED+str(e)+bcolors.ENDC)
        print('-------------------------------')
        return False, 'The Data export Path is invalid'

    #if data export path is valid...continue
    else:
        #check that project number is also valid
        if projectNumberValid:
            #check if the project number is actually a number
            # try:
            #     projectNumber = int(projectNumber)
            # except Exception as e:
            #     print('-------------------------------')
            #     print(bcolors.FAIL+'ERROR'+bcolors.ENDC)
            #     print(bcolors.RED+'The Project Number provided is invalid.'+bcolors.ENDC)
            #     print(bcolors.RED+str(e)+bcolors.ENDC)
            #     print('-------------------------------')
            #     return False
        
            projectNumber = str(projectNumber)
            print('.')
            print('.')
            print('.')
            print('-------------------------------')
            print('Setup')
            print('-------------------------------')
            print('Account Number->',bcolors.HEADER+accountNumber+bcolors.ENDC)
            print('Project Number->',bcolors.HEADER+projectNumber+bcolors.ENDC)
            print('Project Code--->',bcolors.HEADER+projectCode+bcolors.ENDC)
            print('-------------------------------')
            print('.')
            print('.')
            print('.')

            #create folders if they do not exist
            create_folders(projectNumber)

            #####THE BELOW IS FOR CARDED DATA
            #move the carded data to wincross and TAB        
            carded_moved = move_files.move_carded(projectCode,projectNumber,accountNumber)
            
            #move open end file to TAB and rename
            oe_moved = move_files.move_oe(projectCode,projectNumber,accountNumber)

            #####THE BELOW IS FOR OE DATA CLEANUP
            #perform file cleanup
            excel_work(projectCode,projectNumber,accountNumber)
            
            #Prep email
            if(carded_moved and oe_moved and excel_work):
                send_email.send_email(projectNumber)
                return True, 'The process was succesful.'
        else:
            print(bcolors.RED+'The project number:'+bcolors.ENDC,str(projectNumber),bcolors.RED+'is invalid.'+bcolors.ENDC)
