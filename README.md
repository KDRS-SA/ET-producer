# ET-Producer #
A lightweight and simple program designed to combine the "ESSARCH tools for producers" (ETP) and "ESSARCH tools for archivists" (ETA) modules from ESSARCH with the added advantage of running every operation in one step.

## 1. Prerequisites ##

- Requires [7-zip](https://www.7-zip.org/download.html).
- Note: Program has only been tested on windows 10/11 computers.

## 2. Simple Guide ##

1. Run the ET_Producer.exe file and wait for the application window to appear.
2. Press the "Browse Content" button, and select the folder containing every item you wish to add to the dias package's content folder.
2-1. Optionally press "Browse Descriptive Metadata" and/or "Browse Administrative Metadata" to select a folder containing files you wish to add to these folders in the dias package as well.
3. When input folders have been selected press the "Continue" button to proceed to the next page. Here you must fill in all the provided metadata text-input fields (as all of them are mandatory).
3-1. Optionally you may also press the "Import mets.xml Metadata" button at the top of the window, which will open a file explorer which may be used to select a mets.xml file that the program then will use to, in a best effort attempt, autofill some of the input fields. If you use this functionality make sure to doublecheck that all the fields are correct before proceeding.
3-2. Additionally you may also optionally press the "Username: admin" button at the top of the window to change the username of the user creating the dias package (for example to your essarch username).
4. Finally press the "Create Dias Package" button to start the process. This may take some time depending on the amount of content, but when it is done a popup should appear.