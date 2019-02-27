# install packages  ----
#下面指令要在console先跑過，Rmarkdown才不會有問題
library(pipeR) #下面函數有用到pipeR所以要先安裝
library_mul <- function(..., lib.loc = NULL, quietly = FALSE, warn.conflicts = TRUE){
  pkgs <- as.list(substitute(list(...))) %>>% sapply(as.character) %>>% setdiff("list")
  if (any(!pkgs %in% installed.packages()))
    install.packages(pkgs[!pkgs %in% installed.packages()],repos = "http://cran.us.r-project.org")
  sapply(pkgs, library, character.only = TRUE, lib.loc = lib.loc, quietly = quietly) %>>% invisible
}
#tibble,plotly ReporteRs
library_mul(dplyr,lubridate,officer,magrittr)
#officer: ouput print to word
#magrittr:
#optarse: generate exe


# read_data&set input_output_path----
#timestamp = format(Sys.time(), "%m%d-%H%M")
csv_file_path = file.choose()
csv_dir_path = dirname(csv_file_path)
setwd(csv_dir_path)
if(file.exists("./afterQC")==FALSE)
{dir.create("afterQC") }
#system("tree")
#list.dirs():列出

csv_file_path_output=paste0(csv_dir_path,"/afterQC")
csv_dir_path_output = dirname(csv_file_path_output)
csv_file_name = sub(pattern = "(.*)\\..*$", replacement = "\\1", basename(csv_file_path))
data = read.csv(csv_file_path, header=TRUE, sep=",",stringsAsFactors=FALSE)
output_file_path = file.path(csv_file_path_output, paste(csv_file_name,'_afterQC.txt', sep=""))
output_file_path_word = file.path(csv_file_path_output, paste(csv_file_name,'_QCsummary.docx', sep=""))
output_file_path_csv = file.path(csv_file_path_output, paste(csv_file_name,'_QC.csv', sep=""))

#把值貼上去
#cat("### File\n", csv_file_path, file=output_file_path, sep="\n", append=FALSE)


# Timestamp  ----
splitname=unlist(strsplit(csv_file_name, "_"))
#delect last 5 charaacters
start_time=ymd(substr(splitname[3],1,nchar(splitname[3])-5))
end_time=ymd(substr(splitname[4],1,nchar(splitname[4])-5))+hours(24)
#It is only when converting to POSIXct happens that the timezone offset to UTC (six hours for me) enters:
#How to get the beginning of the day in POSIXct, by format
start_time<-format(as.POSIXlt(start_time), "%Y-%m-%d %H:%M")
end_time<-format(as.POSIXlt(end_time), "%Y-%m-%d %H:%M")
ll<-seq.POSIXt(as.POSIXct(start_time),as.POSIXct(end_time),by="10 min")
ll<-ll[1:length(ll)-1]
data_result<-data.frame(time=as.character(format(as.POSIXct(ll), "%Y-%m-%d %H:%M")),stringsAsFactors=FALSE)
data_result$StationID=data[1,"StationID"]
data_result$SensorID=data[1,"SensorID"]

# rep_data ----
#data<-data[c(1,1,1,1,1,2,2,2,3:nrow(data)),]
#data[1,"Rainfall"]=10
#data[7,"Rainfall"]=2
data.overlap<-data %>% 
  group_by(rtime) %>% 
  summarise(n=n()) %>% 
  filter(n>=2)

sum(data.overlap$n)

overlap_set<-data.overlap$rtime

#需確認一下不同重複時間點，雨量值不同的做法?

if(nrow(data.overlap)>0) #Check if a data frame is empty.
{
  print_data_overlap<-data %>% 
    filter(rtime %in% overlap_set) %>%
    group_by(rtime,Rainfall) %>% 
    summarise(n=n()) 
}


data<-data %>%
  distinct(rtime,Rainfall,.keep_all = TRUE)

#wrod output
bold_face <- shortcuts$fp_bold(font.size = 30)
bold_face_small <- shortcuts$fp_bold(font.size = 18)
bold_redface <- update(bold_face, color = "red")
fpar_ <- fpar(ftext(paste0("重複資料筆數=",sum(data.overlap$n),"\n"), prop = bold_face))
#read_docx() reate an R object representing a Word document
doc <- read_docx()%>% body_add_fpar(fpar_)  
print(doc, target = output_file_path_word)

doc <- doc %>%
  body_add_table(print_data_overlap, style = "table_template")
print(doc, target = output_file_path_word )


# combind_data ----
#cannot join a POSIXct object with an object that is not a POSIXct object
#data$rtime<-as.POSIXct(data$rtime)
data$rtime<-format(as.POSIXct(data$rtime), "%Y-%m-%d %H:%M")
# join有錯，有可能是時間的資料格式不對
str(data_result)
str(data)
data_result<-left_join(x = data_result, y = data[,c("rtime","Rainfall")], by = c("time"="rtime"))
length(is.na(data_result$Rainfall)==FALSE)
#check what time loss time
#data_result$Rainfall_F<-factor(data_result$Rainfall, exclude =  "" )
#levels(data_result$Rainfall_F)

# missing_data ----
missing_count<-sum(is.na(data_result$Rainfall)) #6
missing_rate<-round(missing_count/dim(data_result)[1],4)*100
missing_set<-which(is.na(data_result$Rainfall)==TRUE)
data_result[missing_set,]

#word output
bold_face <- shortcuts$fp_bold(font.size = 30)
bold_face_small <- shortcuts$fp_bold(font.size = 18)
bold_redface <- update(bold_face, color = "red")
fpar_1 <- fpar(ftext(paste0("遺失資料筆數=",missing_count), prop = bold_face))
fpar_2 <- fpar(ftext(paste0("遺失資料比例=",missing_rate), prop = bold_face_small))
fpar_3 <- fpar(ftext(paste0("觀測成功率=",(100-missing_rate)), prop = bold_face_small))
#read_docx() reate an R object representing a Word document
#Each section has one fpar,then add new line between sections才能有新的一行
doc <- doc %>% 
  body_add_fpar(fpar_1)  %>% 
  body_add_fpar(fpar_2) %>% 
  body_add_fpar(fpar_3) 
print(doc, target = output_file_path_word)


doc <- doc %>%
  body_add_table(data_result[missing_set,], style = "table_template")
print(doc, target = output_file_path_word )




# output_result ----
colnames(data_result)[which(colnames(data_result)=="time")]<-"rtime"
data_result<-data_result[,c("StationID","SensorID","rtime","Rainfall")]
write.csv(data_result,output_file_path_csv,row.names = FALSE)

