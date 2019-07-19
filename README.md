# checklist
Automatically fills out project checklist with relevant info

Created in Python 3.7.1

Non-standard dependent libraries: editpyxl

I wrote this script to automatically fill in a specific translation project checklist. 
It finds the project number, language pair, and dates that various tasks were done based on the folder names, and inputs them into the checklist along with "〇" to signify that the task was completed. It will do this for any checklist files that it finds in a given directory.
