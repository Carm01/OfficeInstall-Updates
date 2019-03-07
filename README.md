# OfficeInstall-Updates
AutoIT Script processes office updates from *.msp files pulled with vbs

- How to capture office updates:
  - Run Windows updates to update office.
  - Run the CollectUpdates.vbs.
  - Once done a window will open. That Folder is located in your users temp file called "updates".
  - Make sure 7-zip is installed on target as this is the uncompress method.
  - I right click on the updates folder and compress the folder with the updates inside.
  - I use WinRar ( you can use 7-zip if you want you need to change the program code to reflect the changes) to compress the FOLDER, not the files to a file called OfficeUpdates.rar. 
- Depoying updates ( ProcessOfficeUpdates.au3 )
  - Line 39 controls the source and destination for your .rar ( or .7z ) file.
  - Line 44 extracts the contents of the updates.
  - The updates will be processed in order of the oldest "Modified Time" to newest; thus installing the oldest updates to the newest.
