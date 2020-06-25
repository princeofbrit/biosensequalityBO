lag_breakdown<-function(data){
  LagTime=data %>%
    select(C_Biosense_Facility_ID, C_BioSense_ID, Arrived_Date_Time, C_Visit_Date_Time, Message_Date_Time,
           Recorded_Date_Time, Diagnosis_Code)%>% 
    filter(is.na(Diagnosis_Code)==FALSE)%>%
    mutate(Arrived=as.POSIXct(Arrived_Date_Time,format="%Y-%m-%d %H:%M:%S"),
           Visit=as.POSIXct(C_Visit_Date_Time,format="%Y-%m-%d %H:%M:%S"),
           Message=as.POSIXct(Message_Date_Time,format="%Y-%m-%d %H:%M:%S"),
           Record=as.POSIXct(Recorded_Date_Time,format="%Y-%m-%d %H:%M:%S"))%>%
    group_by(C_BioSense_ID)%>%
    slice(which.min(Record))
  
  
  Time_Diff=LagTime%>%
    mutate(lag_Record_Visit=as.numeric(difftime(Record,Visit,units="hours")),
           lag_Message_Record=as.numeric(difftime(Message,Record,units="hours")),
           lag_Arrival_Message=as.numeric(difftime(Arrived,Message,units="hours")),
           lag_Arrival_Visit=as.numeric(difftime(Arrived,Visit,units="hours")) ,
           lag_Message_Visit=as.numeric(difftime(Message,Visit,units="hours")) ,
            category_Record_Visit=cut(lag_Record_Visit, breaks=c(-Inf, 24, 48, Inf), labels=c("<24 Hours","24-48 Hours",">48 Hours")),
           category_Message_Record=cut(lag_Message_Record, breaks=c(-Inf, 24, 48, Inf), labels=c("<24 Hours","24-48 Hours",">48 Hours")),
           category_Arrival_Message=cut(lag_Arrival_Message, breaks=c(-Inf, 24, 48, Inf), labels=c("<24 Hours","24-48 Hours",">48 Hours")),
           category_Arrival_Visit=cut(lag_Arrival_Visit, breaks=c(-Inf, 24, 48, Inf), labels=c("<24 Hours","24-48 Hours",">48 Hours")))
  
  Lag_Record_Visit=Time_Diff %>% ungroup %>%
    group_by(C_Biosense_Facility_ID)%>%
    count(Timeliness=category_Record_Visit)%>%
    mutate(prop_Record_Visit = prop.table(n)*100) %>%
    select(-n) %>%
    complete(C_Biosense_Facility_ID,Timeliness,fill = list(prop_Record_Visit = 0))
  Lag_Message_Record=Time_Diff %>% ungroup %>%
    group_by(C_Biosense_Facility_ID)%>%
    count(Timeliness=category_Message_Record)%>%
    mutate(prop_Message_Record = prop.table(n)*100) %>%
    select(-n)  %>%
    complete(C_Biosense_Facility_ID,Timeliness,fill = list(prop_Message_Record = 0))
  Lag_Arrival_Message=Time_Diff %>% ungroup %>%
    group_by(C_Biosense_Facility_ID)%>%
    count(Timeliness=category_Arrival_Message)%>%
    mutate(prop_Arrival_Message = prop.table(n)*100) %>%
    select(-n) %>%
    complete(C_Biosense_Facility_ID,Timeliness,fill = list(prop_Arrival_Message = 0))
  Lag_Arrival_Visit=Time_Diff %>% ungroup %>%
    group_by(C_Biosense_Facility_ID)%>%
    count(Timeliness=category_Arrival_Visit)%>%
    mutate(prop_Arrival_Visit = prop.table(n)*100) %>%
    select(-n) %>%
    complete(C_Biosense_Facility_ID,Timeliness,fill = list(prop_Arrival_Visit = 0))
Lag_Summary=full_join(full_join(full_join(Lag_Record_Visit,Lag_Message_Record), Lag_Arrival_Message),Lag_Arrival_Visit)
  
  return(Lag_Summary)
  
  
}