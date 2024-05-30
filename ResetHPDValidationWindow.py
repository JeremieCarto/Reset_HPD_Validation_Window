
# Resets the Validation Window at position (0,0) which is the top left corner of the primary/main display.
# Resolves an issue where the position of the window stored in Appsettings refers to a "ghost" monitor.
# J.Robichaud  (https://github.com/JeremieCarto)
# April 3, 2024

import sys        #to interupt and exit the script in case of an error (e.g. sys.exit("Error message"))
import os         #to retrieve the Windows account name (e.g. os.getenv('username'))
import time       #to slightly delay the output messages so that it is easier to read during execution (e.g. time.sleep(seconds))
import subprocess #to acccess the list of running processes in the Windows Task Manager
import xml.etree.ElementTree as ET #This is the XML parser being used

#Function that will look through the list of all Windows processes and look for the specified HPD executable name.
#If the specified HPD app is found, this means that it is currently running and returns True.
def hpd_is_running(hpd_app):
    processes = str(subprocess.check_output('tasklist'))  #Note that the exe name is truncated after 25 characters.
    if hpd_app in processes:
        return True
    else:
        return False

#User is prompted for input
print("")
print("Reset the position of the Validation window for...")
print("  · Source Editor, type \"1\"")
print("  · ENC and Vector, type \"2\"")
print("  · Paper Chart Editor, type \"3\"")

editor = input("Which HPD Editor? ") #store user input in variable "editor"

#User input determines which xml file is assigned to variable "filename"
if editor == "1":
   if hpd_is_running('caris_hpd_source_editor.e'):  #calls the function to check if SE is running. Notice the truncated exe name.
      sys.exit("Error: CARIS Source is running. Please close the software.") 
   else:
      filename = "Source_AppSettings.xml" #This file is used if the app is not running
elif editor == "2":
   if hpd_is_running('caris_hpd_product_editor.'):  #calls the function to check if PE is running. Notice the truncated exe name.
      sys.exit("Error: CARIS ENC and Vector is running. Please close the software.")
   else:
      filename = "ENC_and_Vector_AppSettings.xml" #This file is used if the app is not running
elif editor == "3":
   if hpd_is_running('caris_hpd_paper_chart_edi'):  #calls the function to check if PCE is running. Notice the truncated exe name.
      sys.exit("Error: CARIS Paper Chart is running. Please close the software.")
   else:
      filename = "Paper_Chart_AppSettings.xml" #This file is used if the app is not running
else:
   sys.exit("Error: Valid entries are 1, 2 or 3.") #Error handling of invalid user input

version = input("Which version (\"4.1\" or \"5.0\")? ") #store user input in variable "version"

if version == "4.1" or version == "5.0": #if version is valid then store full path to the xml file in variable "appsettings"
   #concatenation of strings to create path. Windows account name is being retrieved (os.getenv('username'))
   appsettings = "C:\\Users\\" + os.getenv('username') + "\\AppData\\Roaming\\CARIS\\HPD\\" + version + "\\" + filename
else:
   sys.exit("Error: Only HPD versions 4.1 and 5.0 are supported.") #Error handling of invalid user input

tree = ET.parse(appsettings) #XML parser is parsing file stored in variable "appsettings" and result is stored in variable "tree"
root = tree.getroot() #Variable "root" is used to point to the Root Element of the xml document

for XPosition in root.iter('XPosition'): #from the root, iterates through the entire document looking for tag XPosition
   XPosition.set('Value','0') #Set Value to zero for XPosition

for YPosition in root.iter('YPosition'): #from the root, iterates through the entire document looking for tag YPosition
   YPosition.set('Value','0') #Set Value to zero for YPosition

#Note: If there would be multiple XPosition and YPosition tags in the xml document, this script would reset all their values to Zero.
#This is not a problem here as there is only one occurance of each in Appsettings, for the Validation Window only.

print("")
print("Processing " + appsettings) #Information about which file is being processed (including full path)
print("Please wait...")
tree.write(appsettings) #This writes the entire content of the xml document stored in "tree" back to the file
time.sleep(1)
print("")
print("Done!")
print("")