# ET-Producer #
A lightweight and simple program designed to combine the "ESSARCH tools for producers" (ETP) and "ESSARCH tools for archivists" (ETA) modules from ESSARCH with the added advantage of running every operation in one step.

## 1. Prerequisites ##

As far as we know there should be no pre-requisites, although the program has only been tested on windows 10/11 computers.

## 2. Simple Guide ##

1. Run the ET_Producer.exe file, and wait for the terminal and application window to appear. Ignore the terminal for now and go to the application window.
2. Fill in all the provided metadata text-input fields (All fields are mandatory according to the standard, so none should be left empty).
3. Press the "Browse Content" button, and select the folder containing every item you wish to add to the dias package's content folder. Optionally you may do the same with the "Browse Descriptive Metadata" and "Browse Administrative Metadata" buttons to add files to Descriptive Metadata and Administrative metadata folders respectively.
4. Press the "Create Dias Package" button to start the process. While the process is running you can navigate to the terminal window where an update is posted at the start of each internal step.
5. A message should pop-up informing you when the process is finished. When it does, navigate to the folder where the ET_Producer.exe file is located. There will be a new folder here now, which is equivalent to the output of ETA and is ready to be transferred to EPP.

## 3. Importing metadata from mets.xml file ##
1. By pressing the "Import mets.xml Metadata (optional)" button in the GUI that shows up when running the ET_Producer.exe program, a file explorer will show up. Use this to navigate to the xml file containing the mets structure, and press open. The program will then, in a best effort attempt, autofill some of the input fields with the data of the mets xml-file. Make sure to doublecheck that all fields are correct before proceeding.
