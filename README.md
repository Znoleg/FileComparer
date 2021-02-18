# FileComparer
____
## FileComparer is a utility to compare contents of two files.
____
### Currently the highest implementation level is __Level 3__.
#### The solution is divided into 3 project levels to see the differences and progress in creating the program.
#### ***All publish .exe files are located at Publish/ folder***.
#### __How to use:__ When the program starts, it will ask you for the path to the original file and the modified file. Provide a files in .txt/.doc/.docx/.pdf format. Files formats must be the same (can be changed later). If you didn't get an error after some delay program will output files differences in console. In case of level3 it will also create Log file in /logs folder of root directory where you can see string changes in details. It can also contain information about errors if they occurred.
#### Note that you can drag&drop text files to .exe if you're using Level 3 implementation. Also after use it will create a new log file into /logs folder (using NLog library).
###### You can see the details of the program implementation and its operation algorithm in the comments to the code.
###### Used libraries: "microsoft.office.interop.word" for word file reading, "IText7" for PDF reading.
