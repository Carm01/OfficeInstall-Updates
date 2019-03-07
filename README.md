# OfficeInstall-Updates
AutoIT Script processes office updates from *.msp files pulled with vbs

- How to capture office updates:
  - Run Windows updates to update office
  - Run the CollectUpdates.vbs
  - Once done a window will open. That Folder is located in your users temp file called "updates"
  - Make sure 7-zip is installed on target as this is the uncompress method. 
  - I use WinRar ( you can use 7-zip if you want you need to change the program code to reflect the changes) to compress the FOLDER, not the files to a file called OfficeUpdates.rar. 
  - Lines 36-46 control the sorice and target locatios for the rar file
