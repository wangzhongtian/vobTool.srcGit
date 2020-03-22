# a standalone file-based version log and restore software
keyword:
  IronPython,winform  GUI,windows(.Net Framework),linux(mono),version log and restore

fucntions:
  1 Backup the user specified files or folders in vob DB ;Only the altered files are backuped in VOB DB.
  2 Restore requested version of files from  vob DB ; Only the newer files in VOB DB are copyed to target folder.

Use cases: 
  1 User can backup or log the files to vob DB, copy the vob DB to USB or put on internet ,and restore it at another computer. In this procedure,only the altered file would be copyed.
  2 User can backup files ,No data loss will happen at work .
  
information:
  The python(NOT IRONPYTHON) + Tk version will be provided on request.
  
  
Folders:
  filesafe/src: main python files.a02-xx fil is the main python file.
  filesafe/compiletool: compile the python file to binary executable ,if you need.
  
  
