#' Write NSSP BioSense Platform Data Quality Summary Reports for One Facility
#'
#' @description
#' This function is a lightweight version of the `write_reports` function. It will generate summary and example workbooks, but only for
#' one specified facility. The first summary workbook shows percents and counts of nulls and invalids, while the examples workbook
#' generates de tailed information on records and visits that are null or invalid.
#'
#' @param username Your BioSense username, as a string. This is the same username you may use to log into RStudio or Adminer.
#' @param password Your BioSense password, as a string. This is the same password you may use to log into RStudio or Adminer.
#' @param table The table that you want to retrieve the data from, as a string.
#' @param mft The MFT (master facilities table) from where the facility name will be retrieved, as a string.
#' @param start The start date time that you wish to begin pulling data from, as a string.
#' @param end The end data time that you wish to stop pulling data from, as a string.
#' @param facility The C_Biosense_Facility_ID for the facility that you wish to generate and write the report for.
#' @param directory The directory where you would like to write the reports to (i.e., "~/Documents/MyReports"), as a string.
#' @param nexamples An integer number of examples you would like for each type of invalid or null field in the examples workbooks for each facility.
#' This defaults to 0, which will not generate these example workbooks.
#' @import dplyr
#' @import tidyr
#' @import openxlsx
#' @import RODBC
#' @import emayili
#' @importFrom stringr str_replace_all
#' @export
write_facility_report <- function(username, password, table, mft, start, end, facility, directory="",field=NA, exclude=NA, optional=TRUE, email=FALSE, sender=NA,receiver=NA,email_password=NA, name=NA,title=NA, phone=NA, inline=F ) {
  
  if (as.POSIXct(start)>as.POSIXct(end)){
    print("Error: Start time after end time")
  } else {  
    start1=as.POSIXct(start)
    end1=as.POSIXct(end)
    channel <- odbcConnect("BioSense_Platform", paste0("BIOSENSE\\", username), password) # open channel
    data <- sqlQuery(
      channel,
      paste0("SELECT * FROM ", table, " WHERE C_Visit_Date_Time >= '", start1, "' AND C_Visit_Date_Time <= '", end1, "' AND C_Biosense_Facility_ID = ", facility) # create sql query
      , as.is=TRUE)
    
    if (nrow(data) == 0) stop("The query yielded no data.")
    name <- as.character(unlist(unname(c(sqlQuery(channel, paste0("SELECT Facility_Name FROM ", mft, " WHERE C_Biosense_Facility_ID = ", facility)))))) # get name from mft

    odbcCloseAll() # close connection
  # get hl7 values
  data("hl7_values", envir=environment())
  hl7_values$Field <- as.character(hl7_values$Field)
  
  # get facility-level state summary of required nulls
  req_nulls <- get_req_nulls(data) %>%
    select(-c(C_Biosense_Facility_ID)) %>%
    gather(Field, Value, 4:ncol(.)) %>%
    spread(Measure, Value) %>%
    right_join(hl7_values, ., by = "Field")%>%
    select(-c(Feed_Name,Sending_Application))
  # get facility-level state summary of optional nulls
  opt_nulls <- get_opt_nulls(data) %>%
    select(-c(C_Biosense_Facility_ID)) %>%
    gather(Field, Value, 4:ncol(.)) %>%
    spread(Measure, Value) %>%
    right_join(hl7_values, ., by = "Field")%>%
    select(-c(Feed_Name,Sending_Application))
  # get facility-level state summary of invalids
  invalids <- get_all_invalids(data) %>%
    select(-c(C_Biosense_Facility_ID)) %>%
    gather(Field, Value, 4:ncol(.)) %>%
    spread(Measure, Value) %>%
    right_join(hl7_values, ., by = "Field")%>%
    select(-c(Feed_Name,Sending_Application))

  # get overall complete by merging req_null and opt_null
  overall_complete<- full_join(req_nulls,opt_nulls) %>%
    mutate(Percent_Complete=100-Percent) %>%
    select(-c(Count,Percent))
  #create a table of overall complete% and overall valid
  overall<-full_join(overall_complete,invalids) %>%
    mutate(Percent_Valid=100-Percent) %>%
    select(-c(Count,Percent))

  nrow1=length(overall$Field)
  #add warnings
  for (i in 1:nrow1){
    if(is.na(overall$Percent_Valid[i]) & is.na(overall$Percent_Complete[i])) {
        overall$Warning[i]="Warning: Percent Complete and Percent Valid Missing"
    } else if (is.na(overall$Percent_Complete[i])& overall$Percent_Valid[i]<90){
      overall$Warning[i]="Warning: Percent Valid Under 90"
    } else if (is.na(overall$Percent_Valid[i])& overall$Percent_Complete[i]<90){
      overall$Warning[i]="Warning: Percent Complete Under 90"
    }  else if(overall$Percent_Valid[i]<90 & overall$Percent_Complete[i]<90) {
      overall$Warning[i]="Warning: Percent Complete and Percent Valid Under 90"
    } else {
      overall$Warning[i]=NA
    }
  }
  if (is.na(overall$Warning[which(overall$Field=='Age_Reported')])==F & is.na(overall$Warning[which(overall$Field=='Birth_Date_Time')])==T){
    overall$Warning[which(overall$Field=='Age_Reported')]=NA
  }else if (is.na(overall$Warning[which(overall$Field=='Age_Reported')])==T & is.na(overall$Warning[which(overall$Field=='Birth_Date_Time')])==F){
    overall$Warning[which(overall$Field=='Birth_Date_Time')]=NA
  }
  #Declare optional field
  overall$Optional=ifelse(overall$Field %in% opt_nulls$Field,"Optional", "Required" )
  #Remove optional if optional=False
  if (optional==F){
  overall<-overall %>%
    filter(overall$Optional=='Required')
  }
  #select the field needed
  if (is.character(field)==T){
    field1=paste(field,collapse="|")
    overall<-overall%>%
      filter(grepl(field1,overall$Field, ignore.case = T))
  }
  #exclude the select field
  if (is.character(exclude)==T){
    exclude1=paste(exclude,collapse="|")
    overall<-overall%>%
      filter(!grepl(exclude1,overall$Field, ignore.case = T))
  }
  nrow=length(overall$Field)
  # getting first and last visit date times
  vmin <- min(as.character(data$C_Visit_Date_Time))
  vmax <- max(as.character(data$C_Visit_Date_Time))
  amin <- min(as.character(data$Arrived_Date_Time))
  amax <- max(as.character(data$Arrived_Date_Time))
  Lag<-data.frame(
    HL7=c("EVN-2.1","MSH-7.1,EVN-2.1","MSH-7.1",""),
    Lag_Name=c("Record_Visit","Message_Record","Arrival_Message","Arrival_Visit"),
    Lag_Between=c("Record_Date_Time-Visit_Date_Time","Message_Date_Time-Record_Date_Time","Arrival_Date_Time-Message_Date_Time","Arrival_Date_Time-Visit_Date_Time"),
    Group1=t(Lag_Summary[1,3:6]),
    Group2=t(Lag_Summary[2,3:6]),
    Group3=t(Lag_Summary[3,3:6])
  )
  colnames(Lag)[4:6]=c("<24 Hr","24-48 Hr", ">48 Hr")
  ##create overall powerpoint
  wb <- createWorkbook()
  sheet1<- addWorksheet(wb, "Summary")
  writeDataTable(wb, sheet1, overall, firstColumn=TRUE, headerStyle=hs, bandedRows=TRUE) # write Completeness to table
  setColWidths(wb, sheet1, 1:ncol(overall), "auto") # format sheet
  freezePane(wb, sheet1, firstActiveRow=2) # format sheet
  writeDataTable(wb,sheet1,Lag,startCol=1,startRow=nrow+3, headerStyle=hs, colNames=TRUE,rowNames=FALSE,firstColumn=TRUE) #write Timeliness to table

  ##colorcode sheet
  negStyle <- createStyle(fontColour = "#000000", bgFill = "#FFC7CE")
  posStyle <- createStyle(fontColour = "#000000", bgFill = "#C6EFCE")
  midStyle <- createStyle(fontColour = "#000000", bgFill = "#FFFF00")

  naStyle <- createStyle(bgFill = "#808080")

 
  conditionalFormatting(wb, sheet1, cols=3, rows=2:(nrow+1), rule="$C2<90", style = negStyle)
  conditionalFormatting(wb, sheet1, cols=3, rows=2:(nrow+1), rule="$C2>=90", style = posStyle) 
  conditionalFormatting(wb, sheet1, cols=3, rows=2:(nrow+1), rule="ISBLANK($C2)=TRUE", style = naStyle)
  conditionalFormatting(wb, sheet1, cols=4, rows=2:(nrow+1), rule="$D2<90", style = negStyle)
  conditionalFormatting(wb, sheet1, cols=4, rows=2:(nrow+1), rule="$D2>=90", style = posStyle)
conditionalFormatting(wb, sheet1, cols=4, rows=2:(nrow+1), rule="ISBLANK($D2)=TRUE", style = naStyle)
addStyle(wb, sheet1, cols=4, rows=(nrow+4):(nrow+7), style = posStyle)
addStyle(wb, sheet1, cols=5, rows=(nrow+4):(nrow+7), style = midStyle)
addStyle(wb, sheet1, cols=6, rows=(nrow+4):(nrow+7), style = negStyle)
  # write sheet
  filename <- str_replace_all(name, "[^[a-zA-z\\s0-9]]", "") %>% # get rid of punctuation from faciltiy name
    str_replace_all("[\\s]", "_") # replace spaces with underscores
  saveWorkbook(wb, paste0(directory, "/", filename, "_Overall.xlsx"), overwrite=TRUE)
  if (email==T){
 #compose email message
  warningcount=which(!is.na(overall$Warning))
  nwarning= length(warningcount)
  subject= paste(gsub('_',' ',filename),"Report")
 bodytext= paste("<p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>All,</p>
<p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>&nbsp;</p>
<p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>Attached is the data quality report I ran for ",gsub('_',' ',filename),". I have summarized some of the issues I noticed with the data from" , start, " to", end,". Please don&rsquo;t hesitate to contact me with any questions.</p>
<p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><strong><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 14pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'>&nbsp;</span></strong></p>")
  if (inline==T){
  if (nwarning==0){
    bodytext=paste(bodytext,"
                   <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>All the fields are complete and valid, keep up the good work!</p>")
  } else {
    
    for (j in 1:nwarning){
      if(overall$Warning[warningcount[j]]=="Warning: Percent Complete Under 90"){
        bodytext=paste(bodytext,"
                       <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>", overall$Percent_Complete[warningcount[j]],"% of", overall$Field[warningcount[j]],"is complete. The HL7 code is",overall$HL7[warningcount[j]] ,"</p>")
      }
      if(overall$Warning[warningcount[j]]=="Warning: Percent Valid Under 90"){
        bodytext=paste(bodytext,"
                       <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>", overall$Percent_Valid[warningcount[j]],"% of", overall$Field[warningcount[j]],"is valid. The HL7 code is",overall$HL7[warningcount[j]] ,"</p>")
      }
      if(overall$Warning[warningcount[j]]=="Warning: Percent Complete and Percent Valid Under 90"){
        bodytext=paste(bodytext,"
                       <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>", overall$Percent_Complete[warningcount[j]],"% of", overall$Field[warningcount[j]],"is complete.", overall$Percent_Valid[warningcount[j]],"% of", overall$Field[warningcount[j]],"is valid. The HL7 code is",overall$HL7[warningcount[j]] ,"</p>")
      }
      if(overall$Warning[warningcount[j]]=="Warning: Percent Complete and Percent Valid Missing"){
        bodytext=paste(bodytext,"
                       <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>", overall$Field[warningcount[j]],"has missing % Complete and % Valid. The HL7 code is",overall$HL7[warningcount[j]] ,"</p>")
      }
    }
    }
  }

bodytext=paste(bodytext,"<p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><strong><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 14pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'>&nbsp;</span></strong></p>
<p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><strong><em><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 14pt; line-height: inherit; font-family: 'Albany AMT'; vertical-align: baseline; color: inherit;'>Jeanne Jones</span></em></strong></p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>785-296-6339</p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'>Onboarding Coordinator, Syndromic Surveillance</p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 10pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'><a data-auth='NotApplicable' href='mailto:jejones@kdheks.gov' rel='noopener noreferrer' style='margin: 0px; padding: 0px; border: 0px; font: inherit; vertical-align: baseline;' target='_blank'><span style='margin: 0px; padding: 0px; border: 0px; font: inherit; vertical-align: baseline; color: blue;'>jeanne.jones@ks.gov</span></a></span></p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 10pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'><a data-auth='NotApplicable' href='mailto:kdhesys@kdheks.gov' rel='noopener noreferrer' style='margin: 0px; padding: 0px; border: 0px; font: inherit; vertical-align: baseline;' target='_blank'><span style='margin: 0px; padding: 0px; border: 0px; font: inherit; vertical-align: baseline; color: blue;'>kdhe.syndromic@ks.gov</span></a></span></p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 10pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'>KDHE-DPH-BEPHI-PHI-VSDA</span></p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 10pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'>1000 S Jackson Street, STE 130</span></p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 10pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'>Topeka, KS &nbsp;66612</span></p>
         <p style='color: rgb(50, 49, 48); font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; font-size: 11pt; font-family: Calibri, sans-serif; margin: 0px;'><span style='margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; font-size: 10pt; line-height: inherit; font-family: inherit; vertical-align: baseline; color: inherit;'>Kansas Meaningful Use</span></p>
")

    email <- envelope() %>%
      from(sender) %>%
      to(receiver) %>%
      subject(subject) %>%
      html(bodytext) %>%
      attachment(path=paste0(directory, "/", filename, "_Overall.xlsx"))
    
    smtp <- server(host = "smtp.office365.com",
                   port = 587,
                   username = sender,
                   password = email_password)
    
    smtp(email, verbose = TRUE)
  }
}
}

