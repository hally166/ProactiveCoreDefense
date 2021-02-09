#ProactiveCoreDefense script for checking user data quality autonomously
#v2.3 (AUG2020)
#Christopher Hall, EMBL ROME Flow Cytometry Facility, christopher.hall@embl.it
#GPL-3.0

### Things for you to change ###
csvOUT<- "Q:/User_QC/" # the output folder for the results
databaseLocation <- "D:/Users/Public/Documents/Life Technologies/AttuneNxT/Userdata" # location of the fcs files
instrumentName <- "Attune"
threshold <- 10 # you care about files with this percentage cells removed or more
# Lines 54 and 55 need changing depending on cytometer.  Use 55 for BD and 54 for most other instruments. 
# Line 85 includes a command to exclude the performance tracking files.  Change this to match your own or just ignore using #

### Load the required libraries ### Run the startup script in the readme the first time you deploy this ###

library(flowCut)
library(openxlsx)

### FUNCTIONS ### DONT CHANGE UNLESS ADAPTING FOR A DIFFERNT MACHINE ###

#function to load the FCS files and transform them ready for plotting
loadNtransform <- function(ff){
  ff<-tryCatch(read.FCS(ff), error = function(e){read.FCS(ff, emptyValue = FALSE)})
  x<-autovect_verbose(ff)
  biexp  <- biexponentialTransform("myTransform")
  ff<-transform(ff,transformList(x,biexp))
  return(ff)
}

#function to exclude FSC, SSC, and time from transformation - add your own to the grep line is needed
autovect_verbose<- function(ff){
  c<- data.frame(ff@parameters@data)
  d<- grep("FSC|SSC|Time|-W", c$name, invert = TRUE, value = TRUE)
  return(unname(d))
}

#function to perform the flowCut procedure and return the results
flowCut_data<- function(ff){
  outputName<-make.names(paste(toString(keyword(ff)["$DATE"][[1]]),toString(keyword(ff)["$ETIM"][[1]]),keyword(ff)["$FIL"][[1]],sep="_"))
  res_flowCut <- flowCut(ff, FileID = outputName, Directory = paste0(csvOUT,instrumentName,"\\\\",toString(format(Sys.time(), "%Y%b"))))
  if (res_flowCut$data[17] == "F"){
    Link=gsub(".fcs","",gsub("/","\\\\",gsub(" ","_",paste0(csvOUT,instrumentName,"/",toString(format(Sys.time(), "%Y%b")),"/",outputName,"_Flagged_",as.character(res_flowCut$data["Is it monotonically increasing in time", ]),
                                                            as.character(res_flowCut$data["Continuous - Pass", ]),
                                                            as.character(res_flowCut$data["Mean of % - Pass", ]),
                                                            as.character(res_flowCut$data["Max of % - Pass", ]),".png")))) #This line is soooo ugly
  } else {Link=""}
  results_df<-data.frame(
    UniqueID=as.character(outputName),
    PercentEventsRemoved=res_flowCut$data[13],
    WorstChannel=res_flowCut$data[12],
    fcsFileName=keyword(res_flowCut$frame)["$FIL"][[1]],
    Date=toString(keyword(res_flowCut$frame)["$DATE"][[1]]),
    Time=toString(keyword(res_flowCut$frame)["$ETIM"][[1]]),
    Operator=keyword(res_flowCut$frame)["$OP"][[1]], #Attune and Cytoflex
    #Operator=keyword(res_flowCut$frame)["EXPORT USER NAME"][[1]], #BD
    Events=keyword(res_flowCut$frame)["$TOT"][[1]],
    Link = Link,
    stringsAsFactors = FALSE
  )
}

### RUN THE SCRIPT FROM HERE ###

#Load the previous data
tryCatch(prev_data_summary<-read.xlsx(paste0(csvOUT, instrumentName,".xlsx"), sheet = "Summary"), error = function(e){
  prev_data_summary<<-data.frame()
})
tryCatch(prev_data_details<-read.xlsx(paste0(csvOUT, instrumentName,".xlsx"), sheet = "Details"), error = function(e){
  prev_data_details<<-data.frame()
})

#Check to see when the script was last run
if (file.exists(paste0(csvOUT,instrumentName,".xlsx"))) {
  LatestDate<-as.Date(tail(prev_data_summary$DateTested, n=1),"%a %d %b %Y")
} else { LatestDate<- Sys.Date()-7
}

#run the functions and generate the data
inputFiles<-list.files(databaseLocation, full.names = TRUE, ignore.case = TRUE, pattern = ".fcs", recursive = TRUE)
fileinfo<- file.info(inputFiles)
time<-as.POSIXct(LatestDate, format="%Y/%m/%d")+1
files2cut<-fileinfo[fileinfo$mtime>time,]
files2cut<-as.list(row.names(files2cut))
files2cut_noPerf <- files2cut
#files2cut_noPerf <- files2cut[!grepl("PerformanceTestResults", files2cut)] #use this line to use REGEX to exclude certain files, like performace checks

flowCut_results = data.frame()
for (file in files2cut_noPerf) {
  try({flowCut_results <- rbind(flowCut_results,flowCut_data(loadNtransform(file)))},silent = TRUE)
}

if (dim(flowCut_results)[1] == 0) {
  quit(save="no")
}

#Sumarise today's data
command <- paste0("\"$FSO = New-Object -ComObject Scripting.FileSystemObject ; $FSO.GetFolder(\'",databaseLocation, "\').Size")
summaryData <- data.frame(DateTested=toString(format(Sys.time(), "%a %d %b %Y")),
                          DatabaseSizeGb=as.numeric(system2("powershell", args = command, stdout = TRUE))/1000000000,
                          TotalAcquisitions=nrow(flowCut_results),
                          TotalEvents=sum(as.numeric(as.character(unlist(flowCut_results['Events'])))),
                          TotalUsers=sapply(flowCut_results['Operator'], function(x) length(unique(x))),
                          NumberOfBadAcquisitions=sum(as.numeric(as.character(unlist(flowCut_results['PercentEventsRemoved']))) > threshold, na.rm = TRUE),
                          stringsAsFactors=FALSE)

### SAVE THE DATA ###

prev_data_summary<-rbind(prev_data_summary,summaryData)
prev_data_details<-rbind(prev_data_details,flowCut_results)
class(prev_data_details$Link) <- "hyperlink"
wb <- createWorkbook()
addWorksheet(wb, sheetName = "Summary", gridLines = FALSE)
addWorksheet(wb, sheetName = "Details", gridLines = FALSE)
writeDataTable(wb, sheet = "Summary", x = prev_data_summary, colNames = TRUE, rowNames = FALSE,tableStyle = "TableStyleLight9")
writeDataTable(wb, sheet = "Details", x = prev_data_details,colNames = TRUE, rowNames = FALSE,tableStyle = "TableStyleLight9")
setColWidths(wb, sheet = "Summary", cols = 1:10, widths = "auto")
setColWidths(wb, sheet = "Details", cols = 1:10, widths = "auto")
saveWorkbook(wb, paste0(csvOUT, instrumentName,".xlsx"), overwrite = TRUE)
