#!/usr/bin/env Rscript
# This script will only work properly on Microsoft Windows systems with R installed.
# For drag-and-drop or command-line runs, RScript.exe needs to be in a directory of your %PATH%

### Set up default filenames (for Matt's test data set).
d_f <- "P:/Win8Usr/mpagel/Downloads/SC Lab Vemco files-20181015T224436Z-001/Detections_from_SCLabDB2018_all.csv"
sl_f <- "P:/Win8Usr/mpagel/Documents/deploymentSlopes_NOAA.csv"
o_f <- "R_output.csv"

totTime <- proc.time() # keep track of how long things are taking

if (is.null(d_f) || !file.exists(d_f))  {d_f <- NULL}
if (is.null(sl_f) || !file.exists(sl_f)){sl_f <- NULL}

### set up some basic functions

pp<-function(...) {
  # print(paste0(...))
  message(...)
}
warn<-function(...) {
  message("WARNING: ",...)
}

goodLibraryPath <- function() {
  dox <- normalizePath(readRegistry("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders","HCU",maxdepth=1)$Personal,winslash="/",mustWork=F) 
  # Windows doesn't approve of using this registry key folder, but they're not exactly going to take it away any time soon, as many programs depend on values here.
  # The alternate would be to hook into some .NET function calls, which wouldn't be simple to code and would also require .NET framework installed on target systems.
  # Another alternative would just assume that this is located at c:\users\<currentUserName>\Documents\, which is not true for older Windows Systems nor customized systems like mine;)
  
  if (!is.null(dox)) { dox <- paste0(dox,"/R/win-library/x.y") }
  options(repos=structure(c(CRAN="https://cran.cnr.berkeley.edu/")))
  
  lps <- c(dox, paste0(path.expand("~"),"/R/win-library/x.y"), tempdir(), getwd(), .libPaths())
  lps <- lapply(lps, normalizePath, winslash="/",mustWork=F)
  # message("Initially LIBPATH=",paste(lps,collapse=";"))
  lps2 <- lapply(lps, grep, pattern="/x.y", value=TRUE) # grab the paths with generic R version filler only
  rv <- paste(unlist(strsplit(toString(getRversion()),".",fixed=TRUE)[[1]])[1:2],sep=".",collapse=".")
  message("You are running R version: ",rv)
  # substr(as.character(getRversion()),1,3) will not work if subversion exceeds 9
  lps2 <- lapply(lps2, function(x) { sub("/x.y",paste0("/",rv),x,fixed=TRUE) }) # change x.y to e.g. 3.5
  lps <- Filter(file.exists,unlist(c(lps2,lps)))
  # lps <- Filter(file.exists,unlist(lps))
  .libPaths(new=lps)
  message("Library path set. LIBPATH=",paste(.libPaths(),collapse=";"))
}

install.load <- function(package.name) {
  message(paste("Attaching package:", package.name))
  if (!require(package.name, warn.conflicts=TRUE, quietly=TRUE, character.only=TRUE)) {
    warn("Package ", package.name, " not found in ", paste0("LIBPATH=",paste(.libPaths(),collapse=";")), " we will attempt download and installation from cran-R.")
    install.packages(package.name, quiet=TRUE)
    if (!library(package.name, character.only=TRUE, logical.return=TRUE, warn.conflicts=TRUE, quietly=TRUE)) {
      message(paste("ERROR: Could not download & install required package:",package.name))
      q(status=1)
    } else {
      message(paste("Downloaded and installed",package.name))
    }
  }
}

ISO2Human<-function(x) {
  ot<-substr(x,1,19)
  substr(ot,11,12)<-" "
  if(anyNA(ot)) ot<-substr(x,1,10)
  return(ot)
}

opt <- {}
if (interactive()) { 
  # WARNING: do not run this as a multiline block. Only as an entire script or the THEN clause lines one at a time 
  #    (aside from this 3-4 lines, the rest of this script can be mass-run).
  opt$datafile <- readline(prompt = "VEMCO data text file name: ")
  opt$outfile <- readline(prompt = "Output file name: ")
  opt$slopefile <- readline(prompt = "slope file name: ")
} else {
  goodLibraryPath()
  install.load('getopt')
  spec <- matrix(c(
    'datafile', 'd',2,'character',
    'outfile',  'o',2,'character',
    'slopefile','s',2,'character',
    'help',     '?',0,'logical'
  ), byrow=TRUE, ncol=4)
  opt <-getopt(spec, debug=TRUE)
}
if (!is.null(opt$datafile)  && nchar(opt$datafile)  > 2 && file.exists(opt$datafile))  {d_f <- opt$datafile}
if (!is.null(opt$outfile)   && nchar(opt$outfile)   > 2)                               {o_f <- opt$outfile}
if (!is.null(opt$slopefile) && nchar(opt$slopefile) > 2 && file.exists(opt$slopefile)) {sl_f<- opt$slopefile}
if (!is.null(opt$help) || is.null(d_f)) {
  cat(getopt(spec, usage=TRUE))
  q(status=1)
}
if (file.exists(o_f)) {
  cp_f <- tempfile(fileext=".bak")
  warn("Output file ",o_f, " already exists. Backing up to ", cp_f)
  file.copy(o_f, cp_f, overwrite=TRUE, copy.date=TRUE)
  warn("Note that the newly created backup file is NOT automatically cleaned out by this script")
}

install.load('data.table')
install.load('fasttime')
pp("Parameters read and libraries loaded @ ",timetaken(totTime))

main <- function() {
  if (!is.null(sl_f)) {
    sl<-fread(sl_f)
    slcn<-colnames(sl)
    nslcn<-tolower(slcn)
    setnames(sl,slcn,nslcn)
    sl_timecols <- c("rx_utc_start", "rx_utc_end", "pc_utc_start", "pc_utc_end")
    sl[,(sl_timecols):=lapply(lapply(.SD,fastPOSIXct,tz="GMT"),setattr,"class","numeric"),.SDcols = sl_timecols]
    if (!("slope" %in% nslcn)) sl[,slope:=(pc_utc_end - pc_utc_start)/(rx_utc_end - rx_utc_start)]
    sl[,c("RTy","RSN","DLDate","Ext"):=tstrsplit(filename,"_",fixed=TRUE)][
      ,rx:=paste(RTy,RSN,sep="-")][
        ,c("RTy","RSN","DLDate","Ext"):=NULL]
  } else {
    sl<-data.table(1)[,`:=`(c("pc_utc_start", "rx_utc_start", "pc_utc_end", "rx_utc_end", "slope"),numeric())][,V1:=NULL][,rx:=""][.0]
  }
  pp("Slope file read @ ",timetaken(totTime))
  
  dat<-fread(d_f,fill=TRUE)
  pp("Data file read @ ",timetaken(totTime))
  
  labls<-c("DT_UTC","Receiver","Tx","TN","TxSN","Data","Units","SN","Lat","Lon")
  setnames(dat,labls)
  dat[,`:=`(TN=NULL,SN=NULL,Lat=NULL,Lon=NULL,TxSN=NULL)]
  dat[,DT_UTC:=setattr(fastPOSIXct(DT_UTC,tz="GMT"),"class","numeric")]
  dat[,duptime:=DT_UTC]
  setkey(dat,Receiver,DT_UTC,duptime)
  setkey(sl,rx,rx_utc_start,rx_utc_end)
  pp("Data tables keyed @ ",timetaken(totTime))
  fo<-foverlaps(dat,sl,type="within",mult="last",nomatch=NA)
  pp("Detections linked to receiver downloads @ ",timetaken(totTime))
  fo[,`:=`(duptime=NULL,Data2="",Units2="")][
    ,DetectDate:=setattr(((DT_UTC-rx_utc_start)*slope) + pc_utc_start - 28799.5,"class",c("POSIXct","POSIXt"))][
      is.na(DetectDate),`:=`(DetectDate=setattr(DT_UTC - 28799.5,"class",c("POSIXct","POSIXt")),Units2="NoDriftCorrection")]
  fo[,c("TxFreq","TxCs","TagID"):=tstrsplit(Tx,"-",fixed=TRUE)][
    ,c("RecTy","VR2SN"):=tstrsplit(Receiver,"-",fixed=TRUE)][
      ,Codespace:=paste(TxFreq,TxCs,sep="-")][
        ,c("TxFreq","TxCs","RecTy","Receiver","vrlid","startvrl","rxid","filename","rx_utc_start","rx_utc_end","pc_utc_start","pc_utc_end","slope","DT_UTC","Tx"):=NULL]
  setcolorder(fo,c("TagID","Codespace","DetectDate","VR2SN","Data","Units","Data2","Units2"))
  pp("Timestamps adjusted to PST (UTC-08:00) @ ",timetaken(totTime))
  
  tf<-tempfile()
  fwrite(fo,file=tf)
  fo<-fread(tf)
  fo[,DetectDate:=ISO2Human(DetectDate)]
  unlink(tf)
  pp("Coerced to string @ ",timetaken(totTime))
  tryCatch(
    fwrite(fo, file = o_f, quote = FALSE), error=function(e) {
      cp_f <- tempfile(pattern="R_output_",fileext=".csv")
      warn("Error [",e,"] occurred writing output file ",o_f)
      warn("Attempting write to ",cp_f, " instead. If this fails, no further rescue will be attempted.")
      try(fwrite(fo, file = cp_f, quote = FALSE))
    })
  pp("Output file written @ ",timetaken(totTime))
}
main()
