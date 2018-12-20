#!/usr/bin/env Rscript
# commandArgs() -> "RStudio"       "--interactive"
if (interactive()) { # warning: do not run this as a multiline block. Only as an entire script or this part disjointed by the following.
  d_f <- "P:/Win8Usr/mpagel/Downloads/SC Lab Vemco files-20181015T224436Z-001/Detections_from_SCLabDB2018_all.csv"
  sl_f <- "P:/Win8Usr/mpagel/Documents/deploymentSlopes_NOAA.csv"
  o_f <- "R_output.csv"
  df2 <- readline(prompt = "VEMCO data text file name: ")
  if (nchar(df2) > 2) {d_f <- df2}
  sl2 <- readline(prompt = "slope file name: ")
  if (nchar(sl2) > 2) {sl_f <- sl2}
  of2 <- readline(prompt = "Output file name: ")
  if (nchar(of2) > 2) {o_f <- of2}
  rm(df2,sl2,of2)
} else {
  ca <- commandArgs() # do something with this...
}
totTime<-proc.time()

install.load <- function(package.name)
{
  if (!require(package.name, character.only=T)) install.packages(package.name)
  library(package.name, character.only=T)
}
pp<-function(...) {
  print(paste0(...))
}
ISO2Human<-function(x) {
  ot<-substr(x,1,19)
  substr(ot,11,12)<-" "
  if(anyNA(ot)) ot<-substr(x,1,10)
  return(ot)
}

install.load('data.table')
install.load('fasttime')
pp("parameters read and libraries loaded: ",timetaken(totTime))

main <- function() {
  sl<-fread(sl_f)
  sl_timecols <- c("Rx_UTC_Start", "Rx_UTC_end", "PC_UTC_Start", "PC_UTC_End")
  sl[,(sl_timecols):=lapply(lapply(.SD,fastPOSIXct,tz="GMT"),setattr,"class","numeric"),.SDcols = sl_timecols]
  sl[,c("RTy","RSN","DLDate","Ext"):=tstrsplit(FileName,"_",fixed=TRUE)][
    ,Rx:=paste(RTy,RSN,sep="-")][
      ,c("RTy","RSN","DLDate","Ext"):=NULL]
  pp("slope File Read: ",timetaken(totTime))
  dat<-fread(d_f,fill=TRUE)
  pp("data Fie Read: ",timetaken(totTime))
  labls<-c("DT_UTC","Receiver","Tx","TN","TxSN","Data1","Units1","SN","Lat","Lon")
  setnames(dat,labls)
  dat[,`:=`(TN=NULL,SN=NULL,Lat=NULL,Lon=NULL,TxSN=NULL)]
  dat[,DT_UTC:=setattr(fastPOSIXct(DT_UTC,tz="GMT"),"class","numeric")]
  dat[,duptime:=DT_UTC]
  setkey(dat,Receiver,DT_UTC,duptime)
  setkey(sl,Rx,Rx_UTC_Start,Rx_UTC_end)
  pp("data tables keyed: ",timetaken(totTime))
  fo<-foverlaps(dat,sl,type="within",mult="last")
  pp("detections linked to receiver downloads: ",timetaken(totTime))
  fo[,`:=`(duptime=NULL,Data2="",Units2="")][
    ,DetectDate:=setattr(((DT_UTC-Rx_UTC_Start)*Slope)+PC_UTC_Start-28799.5,"class",c("POSIXct","POSIXt"))][
      is.na(DetectDate),`:=`(DetectDate=setattr(DT_UTC-28799.5,c("POSIXct","POSIXt")),Units2="NoDriftCorrection")]
  fo[,c("TxFreq","TxCs","TagID"):=tstrsplit(Tx,"-",fixed=TRUE)][
    ,c("RecTy","RecSN"):=tstrsplit(Receiver,"-",fixed=TRUE)][
      ,Codespace:=paste(TxFreq,TxCs,sep="-")][
        ,c("TxFreq","TxCs","RecTy","Receiver","VRLID","StartVRL","RxID","FileName","Rx_UTC_Start","Rx_UTC_end","PC_UTC_Start","PC_UTC_End","Slope","DT_UTC","Tx"):=NULL]
  setcolorder(fo,c("TagID","Codespace","DetectDate","RecSN","Data1","Units1","Data2","Units2"))
  pp("timestamps adjusted to PST (UTC-08:00): ",timetaken(totTime))
  
  tf<-tempfile()
  fwrite(fo,file=tf)
  fo<-fread(tf)
  fo[,DetectDate:=ISO2Human(DetectDate)]
  unlink(tf)
  pp("coerced to string: ",timetaken(totTime))
  fwrite(fo, file = o_f, quote = FALSE)
  pp("output file written: ",timetaken(totTime))
}
main()
