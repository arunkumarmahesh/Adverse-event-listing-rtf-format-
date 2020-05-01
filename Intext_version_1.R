###################################
####### Dynamic Listing AEL001 ####
###################################

####### Reading mulitple libraries in one go

#trace(utils:::unpackPkgZip, quote(Sys.sleep(2)), at = which(grepl("Sys.sleep", body(utils:::unpackPkgZip), fixed = TRUE)))

Libraries <- c("Hmisc", "officer","magrittr", "grid", "gridExtra","gtable", "foreign", "sas7bdat", "dplyr","rtf") 

lapply (Libraries,library, character.only = TRUE)

#library(extrafont)

#windowsFonts(A = windowsFont("Courier New"))

#########################################

data <- read.sas7bdat("C:\\Users\\MAHESAR1\\Desktop\\new\\Dynamic reporting\\adae.sas7bdat")

######## Creating function 

# Print out data from the adverse event dataset

data <- function(data) {
  
  
  #data <- read.sas7bdat(data)
  
  #data = read.sas7bdat("C:\\Users\\MAHESAR1\\Desktop\\new\\Dynamic reporting\\adae.sas7bdat")
  
  data <- subset (data, select = c(CCS, ASR, AESER, AETERM, AEDECOD, AEBODSYS, AESTDTC, AEENDTC, AETOXGR, AERELN,
                                   AEACNOTN, AEACNN, AEOUTN, SAFFL, ASTDY, AENDY))
  
  
  data <- data.frame(data, stringsAsFactors = FALSE)
  
  data$AESTDTC <- as.Date( data$AESTDTC , format = "%Y-%m-%d")
  
  data$AEENDTC <- as.Date( data$AEENDTC, format = "%Y-%m-%d")
  
  data$DURATION <- data$AEENDTC - data$AESTDTC
  
  data$AESTDTC <- as.character( data$AESTDTC)
  
  data$AEENDTC <- as.character( data$AEENDTC)
  
  data$AEENDTC <- ifelse(is.na(data$AEENDTC), "", data$AEENDTC)
  
  data$AENDY <- ifelse(is.na(data$AENDY), "", data$AENDY)
  
  ############################
  ### Sorting data
  ############################
  
  sort <- c("CCS", "AESTDTC", "AEENDTC") ##### select parameter names to sort in the dataset
  
  data <- arrange_(data, .dots = sort)
  
  
  #########################
  #### Filter
  #########################
  
  fil_ter <- quote(SAFFL %in% "Y")  
  
  ## If we want to filter any observation in the parameters we can use this step, here SAFFL is the parameter and "Y" is the observations
  
  data <- filter_(data, .dots = fil_ter) 
  
  #################################
  #### Concatenating parameters
  #################################
  
  PTTERM <- "PTTERM" 
  
  data <- mutate (data, !!PTTERM := paste (AETERM, AEDECOD, AEBODSYS, sep ="/ "))
  
  start_date <- "AEST" 
  
  data <- mutate (data, !!start_date := paste (AESTDTC, ASTDY, sep ="/ "))
  
  End_date <- "AEEND" 
  
  data <- mutate (data, !!End_date := paste (AEENDTC, AENDY, sep ="/ "))
  
  
  ###############################
  ##### selecting variable names
  ################################
  
  column_names1 <- c ("CCS", "ASR", "AESER", "PTTERM" , "AEST", "AEEND", "DURATION", "AETOXGR", "AERELN",
                      "AEACNOTN", "AEACNN", "AEOUTN")
  
  ## In this example we have taken adverse event	as a dataset and passed required parameter names which need to be derived further				
  
  data <- select_ ( data, .dots = column_names1) #### In this step the dataset is build with specified colum names above
  
  ##################################
  #### Renaming the column Names
  ##################################
  
  CCS <- paste0("country/","\n","Subject","\n","identifier", collapse= " ") ### renaming column name to country subject identifier 
  
  data <- rename (data, !!CCS := CCS) ### Concatenating country and subject identifier
  
  ASR <- paste0("Age/","\n","Sex","\n","Race", collapse= " ") 
  
  data <- rename (data, !!ASR := ASR)
  
  PTTERM <- paste0("Reported term/","\n","Preferred term/","\n"," System organ class", collapse= " ") 
  
  data <- rename (data, !!PTTERM := PTTERM)
  
  AEST <- paste0("Start","\n","date/","\n"," day", collapse= " ") 
  
  data <- rename (data, !!AEST := AEST)
  
  AEEND <- paste0("End","\n","date/","\n"," day", collapse= " ") 
  
  data <- rename (data, !!AEEND := AEEND)
  
  AESER <- paste0("Serious","\n"," event", collapse= " ") ## renaming the parameter names as per listings
  
  data <- rename (data, !!AESER := AESER)
  
  AERELN <- paste0("Cau-","\n","sal-", "\n", "ity", collapse= " ") ## renaming the parameter names as per listings
  
  data <- rename (data, !!AERELN := AERELN)
  
  AEACNOTN <- paste0("Action","\n","taken", "\n", "with","\n", "med." , collapse= " ") ## renaming the parameter names as per listings
  
  data <- rename (data, !!AEACNOTN := AEACNOTN)
  
  AEACNN <- paste0("Med","\n","or", "\n", "ther","\n", "taken" , collapse= " ") ## renaming the parameter names as per listings
  
  data <- rename (data, !!AEACNN := AEACNN)
  
  AEOUTN <- paste0("Out-","\n","come", collapse= " ") ## renaming the parameter names as per listings
  
  data <- rename (data, !!AEOUTN := AEOUTN)
  
  DURATION <- paste0("Dur","\n","ation","\n","(days)",  collapse= " ") ## renaming the parameter names as per listings
  
  data <- rename (data, !!DURATION := DURATION)
  
  AETOXGR <- paste0("Toxi","\n","city","\n","Grade",  collapse= " ") ## renaming the parameter names as per listings
  
  data <- rename (data, !!AETOXGR := AETOXGR)
  
  #################################
  #### Printing into PDF document
  #################################
  
  data$CAL <- NA
  
  
  if (nrow(data)>300){ # this condition is static need to change into dynamic when we use this sort of conditions
    
    data$CAL <- rep(seq(1,nrow(data),5),each =5)
    
  }

  #####################################################
  ####### Exporting into word document
  #####################################################
  
  rtf <- RTF("New listing.rtf",   width=11.69,height=8.27,font.size=9,omi=rep(1,1,1,1))
  
  for ( i in unique(data$CAL)){
    
    new <-  data [data$CAL == i , ]
    
    new$CAL <- NULL
    
    #name.width <- max(sapply(names(new), nchar))
   # new <- format(new, justify = "centre")
   
   addHeader.RTF(rtf, title = "\t\t\t\t\t\t\t\t\t Adverse event", subtitle= "\t\t\t\t\t\t\t\t\t Safety population", font.size=9, top = T)
  
   addHeader.RTF(rtf, title = "Subset: <subset id name>", subtitle= "Actual Treatment: XXXXXXXXXX", font.size=9, top = T)
  
   addTable.RTF(rtf,as.data.frame(new),font.size=9,row.names=FALSE,NA.string="-", col.justify='L',header.col.justify='L', col.widths=rep(c(0.9, 0.9, 0.8, 2.0, 1, 1, 0.6, 0.9,
                                                                                                                                           0.6, 0.6, 0.6, 0.6)))
  
   startParagraph.RTF(rtf)
  
   addText.RTF(rtf,paste("\n","- Day is relative to the reference start date. Missing dates are based on imputed dates.\n"))
  
   addText.RTF(rtf,"- Severity: MILD=Mild, MOD=Moderate, SEV=Severe.\n")
  
   addText.RTF(rtf,"- Relationship to study treatment (causality): 1=No, 2=Yes, 3= Yes, investigational treatment, 4= Yes, other study treatment (non-investigational), 5= Yes, both and/or indistinguishable.")
  
   addText.RTF(rtf,"- Action taken: 1=None, 2=Dose adjusted, 3=Temporarily interrupted, 4=Permanently discontinued, 997=Unknown, 999=NA.")
  
   addText.RTF(rtf,"- Outcome: 1=Not recovered/not resolved, 2=recovered/resolved 3=recovering/resolving, 4=recovered/resolved with sequelae","\n", "5=Fatal, 997=Unknown.")
  
   addText.RTF(rtf,"- Medication or therapies taken: 1=No concomitant medication or non-drug therapy, 10=Concomitant medication or non-drug","\n", "therapy.")
  
   endParagraph.RTF(rtf)
   
   
               
   addPageBreak(rtf, width=11.69,height=8.27,font.size=9,omi=rep(1,1,1,1))
   
  }
  
  
  done.RTF(rtf)
  
  #embed_fonts("New listing.rtf", format = "Courier New", outfile = "New listing-embed.rtf")
}



print(data( data1 ))

################################################
#######################
#############

