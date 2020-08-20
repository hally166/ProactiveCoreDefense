# ProactiveCoreDefense
Protect your flow cytometry core from bad data, however that may arise.

Autonomously monitor user sample quality for flow cytometers.

This is an "improved" version of the 'Sample Quality Monitor' repository I developed at Sanger (https://github.com/hally166/SampleQualityMonitor).  
The primary differences are:
* Moved most of the workload into functions to help with portibility accross instruments
* Removed the Sanger specific statistics  
* Moved most of the variables to the first lines to help with setup
* Removed the ability to email the core.  This was very problamatic and insecure.  Now the data is saved to an Excel file and the image output is the default one form flowCut

The idea here is to test user sample quality by checking time vs fluorescence, as we do during analysis, but at the time of acquisition.  This allows the core facility to be proactive helping our users spot problematic experiments and allows us to check for machine issues, such as recurrent blockages. 

It finds the fcs files produced since that last operation and runs them through the R package flowCut (https://github.com/jmeskas/flowCut) which looks for deviations in fluorescence over time.  The script then plots and records the “bad” files and saves the data to a spreadsheet.  

The output currently looks like this.

![example image](/example.PNG)

The big question is; how do you report this to the user?  We are still working on that.

## Instructions
You run the script on the flow cytometer PC and save the data to the local network drive or to the PC.  I prefer the network drive and I make sure that all my PCs map the drives with the same drive letter.

### On the flow cytometer PC.
* Install R
* Open R from the start menu and install the required R packages:
```R
install.packages("devtools")
devtools::install_github("jmeskas/flowCut")
install.packages("openxlsx")
```
You may be asked for some user input, I normally select 'no'.

### On the network or flow cytometer PC
* Download the QC Script R file and ps1 file and save them to the network or local PC
* Open the file and change the options at the top.  It is explained in the script
* Now the hardest part (possibly).  Flow Cytometer manufacturers are incapable of producing consitant fcs files.  One of th eoptions is to record and to count the number of users, whic is very useful. Sadly, the manufacturers record this is different ways {sigh}.  Look for the following section and you will see some parts commented out using the # charactor.  Add or remove these depending on your instruemnt. By default I have left the $OP keyword in, but if you have a BD machine you will need to add a # to the start of this line and remove the one on the line that uses the 'EXPORT USER NAME' keyword. 
```R
#function to perform the flowCut procedure and return the results
flowCut_data<- function(ff){
```
* Save the file and close it.
* Open the ps1 file (PowerShell) and change the R path to where yours is and the R script path to where that is too.  Save the file and close. 
* Go to task manager and add a new daily task that runs the R script though PowerShell.exe with this arguemnt (or similar)
> -ExecutionPolicy ByPass -File Q:\User_QC\QCScriptPwrShell.ps1

### Testing
There are two ways to test that it works; using RStudio or running it though the command line.  To run it in the command line click on 'Start' type 'CMD' and open the command line.  Then navigate to the PowerScrit file and run it 
> PowerShell.exe -ExecutionPolicy ByPass -File Q:\User_QC\QCScriptPwrShell.ps1
If this does not work, try running it in RStudio and do some troubleshooting.  I'll make a FAQ and video later.

You should nd up with an excel spreadsheet with the data including links to the image files and a new folder with the images inside.
