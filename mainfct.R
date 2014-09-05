
#save excel sheet as "unicode text" or "Text (MS-DOS)" with extension *.txt
# enter directory and filename bnelow and then source this script or highlight everything and run.
# the result will be a file with the name <filename.kodebuch> in the original folder.

# directory <- "C:/Dropbox/current/work/DATA STORAGE/kodebuchsys/makeFromXls" #directory in which the excel sheet resides
# filename <- "bsp.txt" #name of the txt file

makeKodebuch <- function(dir=".", filename=NULL, showLog=TRUE) {
	completePathToFile <- paste0(gsub("/+$","", dir),"/",filename)
	if (!file.exists(completePathToFile)) { stop (paste0("The indicated file: ",filename," does not exist in ",dir,".\nPlease check whether path and file name are correct.")) }
  
	rangeFormat <- "^[0-9\\.\\-]+\\-[0-9\\.\\-]+$"
	valueFormat <- "^\\-?[0-9\\.]+\\;\\-?[0-9\\.]+"
	
	################################################
  origtext <- readLines(completePathToFile)[-c(1:4)]
  returnObject <- list()
  for (i in origtext){
    lineclean <- i
    lineclean <- gsub("^[[:space:]]*", "", gsub("[[:space:]]*$", "", gsub("\t[[:space:]]*", "\t", gsub("[[:space:]]*\t", "\t", lineclean))))
    
    parts <- strsplit(gsub("\"","",lineclean), split="\t")[[1]]
    legalvalueslist <- lapply(strsplit(parts[6], split=";")[[1]], function(x) strsplit(x,split="=")[[1]])
    missingvalueslist <- lapply(strsplit(parts[7], split=";")[[1]], function(x) strsplit(x,split="=")[[1]])
  	
  	print(parts[[4]])
  	valuetype <- ifelse(
	    		parts[4] == "Zeichenkette" , "STRING"
	    		, ifelse (length(grep(rangeFormat, parts[4])) > 0, "RANGE"
	    				, "VALUES"))
  		
  		
    newVar <- list(
      varname=toupper(parts[1])
      , varlabel=gsub("^[[:blank:]]*", "", gsub("[[:blank:]]*$", "", parts[2]))
      , varinstruction = parts[3]
      , varlegal=strsplit(parts[4], split=";")[[1]]
      , varmissing=strsplit(parts[5], split=";")[[1]]
      , varlegallabels=legalvalueslist
      , varmissinglabels=missingvalueslist
    	, valuetype=valuetype
    )
    returnObject[[parts[1]]] <- newVar
  }
  kodebuchfile <- gsub("\\..+",".kodebuch", paste0(gsub("/+$","", dir),"/",filename))
  kodebuchfilestringarray <- c()
  for (i in returnObject){
    thisType <- ifelse(length(grep(valueFormat, i$varlegal))>0, "multipleValues", "indivValues")
    if (thisType == "indivValues") {
        casesLeft1 <- setdiff(i$varlegal, unlist(lapply(i$varlegallabels, function(x) x[1])))
        if (length(casesLeft1) > 0 && (length(i$varlegal) > 1 || length(i$varlegallabels) > 1)) {
          for (k in casesLeft1) {i$varlegallabels[[length(i$varlegallabels)+1]] <- c(k,"")}
        }
        casesLeft2 <- setdiff(unlist(lapply(i$varlegallabels, function(x) x[1])), i$varlegal)    
        if (length(casesLeft2) > 0 && (length(i$varlegal) > 1 || length(i$varlegallabels) > 1)) {
          for (k in casesLeft2) {i$varlegal[[length(i$varlegal)+1]] <- k}
        }
    }
    else { 
      if (!grep(rangeFormat, i$varlegal) && !grep("^Zeichenkette$", i$varlegal, ignore.case=TRUE)) { 
        warning("The legal value specifications and the legal value labels of variable ",i$varname," conflict. Please correct.")
      }
    }
  	
    missingsLeft1 <- setdiff(unlist(lapply(i$varmissinglabels, function(x) x[1])), i$varmissing)
    if (length(missingsLeft1) > 0) { 
      warning("Variable ",i$varname," has labels defined for values that are not declared in the missing values. The value", ifelse(length(missingsLeft1)>1,"s","")," ", paste(missingsLeft1, collapse=";")," ", ifelse(length(missingsLeft1)>1,"are","is")," added to missing value specification. If this is undesired, please change ", filename," and run this script again.", sep="")
    }
    missingsLeft2 <- setdiff(i$varmissing, unlist(lapply(i$varmissinglabels, function(x) x[1])))
    if (length(missingsLeft2) > 0) { 
    	addedEntries <- paste0(
    		'[',missingsLeft2, ' ', paste0('"MISSING VALUE ', missingsLeft2,'"]'), collapse='; '
    	)
      warning("Variable ",i$varname," has missing values declared for which there is no corresponding value label. The default label", ifelse(length(missingsLeft2)>1,"s "," "), addedEntries," ", ifelse(length(missingsLeft2)>1,"are","is")," added to the missing value label specification. If this is undesired, please change ", filename," and run this script again.", sep="")
    }

  if (length(missingsLeft1) > 0 && (length(i$varmissing) > 1 || length(i$varmissinglabels) > 1)) {
    for (k in missingsLeft1) {i$varmissing[[length(i$varmissing)+1]] <- k}
  }
  if (length(missingsLeft2) > 0 && (length(i$varmissing) > 1 || length(i$varmissinglabels) > 1)) {
    for (k in missingsLeft2) {i$varmissinglabels[[length(i$varmissinglabels)+1]] <- paste0(k,' "MISSING VALUE ',k,'"')}
  }
  
  legalvaluesexpand <- ifelse(
  	length(i$varlegal) == 1
    , paste0('{',i$varlegal,'} "',i$varlegallabels,'"')
    , paste(
    		sort(unlist(lapply(i$varlegallabels, function(x) paste0(x[1], ' "',x[2],'"')))), collapse="\n"
    	)
  )
  returnObject[[i$varname]]$legalvaluesexpand <- legalvaluesexpand
  
  missingvaluesexpand <- c()
  for (j in i$varmissinglabels) { 
  	missingvaluesexpand <- c(missingvaluesexpand, paste0(paste0(j, collapse=' "'),'"'));
  }
  missingvaluesexpand <- paste(missingvaluesexpand, collapse="\n")
  returnObject[[i$varname]]$missingvaluesexpand <- missingvaluesexpand
  	
  #print(returnObject[[parts[1]]])
  
  kodebuchfilestringarray <- c(kodebuchfilestringarray, 
                               paste(
                                 i$varname
                                 , i$varlabel
                                 , paste0("\"",i$varinstruction,"\"")
                                 , paste0("{",paste(i$varlegal, collapse=";"),"}")
                                 , paste0("{",paste(i$varmissing, collapse=";"), "}")
                                 , legalvaluesexpand
                                 , missingvaluesexpand
                                 , sep="\n"
                               )
  														)
  }
	cat(paste(kodebuchfilestringarray, collapse="\n\n"), file=kodebuchfile, append=FALSE)
	

	### create spss syntax per variable
	sink(file=paste0(gsub("\\.[[:alnum:]]+$", "_perVariable.sps", completePathToFile)))
	cat("* SPSS Syntax to import metadata to an SPSS data file in per variable order. Run this syntax.\n\n")
	for (i in returnObject){
		cat ("VARIABLE LABELS ",i$varname, " '",i$varlabel,"/",i$varinstruction,"'.\n", sep="")
		if ( i$valuetype == "VALUES")  { 
			cat ("VALUE LABELS ",i$varname, " ", gsub("[[:space:]]*\"", "'", gsub("\n"," ", i$legalvaluesexpand)), ".\n", sep="")		}
		cat ("MISSING VALUES  ",i$varname, " (", paste(i$varmissing, collapse=","), ").\n", sep="")
		cat ("VARIABLE ATTRIBUTE  VARIABLES=",i$varname," ATTRIBUTE=MissingDefinition ('",paste(unlist(lapply( i$varmissinglabels, function(x) { paste0(x[1],"=", x[2])} )), collapse=";"),"').\n", sep="")
		cat(ifelse(i$valuetype=="STRING","STRING ","NUMERIC "),i$varname,".\n", sep="")
		cat("\n")
	}
	sink()
	
	### create spss syntax per type
	sink(file=gsub("\\.[[:alnum:]]+$", "_perType.sps", completePathToFile))
	cat("* SPSS Syntax to import metadata to an SPSS data file in per type order. Run this syntax.\n\n")
	for (i in returnObject){
		cat ("VARIABLE LABELS ",i$varname, " '",i$varlabel,"/",i$varinstruction,"'.\n", sep="")
	}
	cat("\n")
	
	#### create spss syntax per variable
	for (i in returnObject)	{ 
		if ( i$valuetype == "VALUES")  { 
			cat ("VALUE LABELS ",i$varname, " ", gsub("[[:space:]]*\"", "'", gsub("\n"," ", i$legalvaluesexpand)), ".\n", sep="")		}
	}
	cat("\n")
	for (i in returnObject) {
		cat ("MISSING VALUES  ",i$varname, " (", paste(i$varmissing, collapse=","), ").\n", sep="")
	}
	sink()
	
	### create R syntax
	sink(file=gsub("\\.[[:alnum:]]+$", "_forR.R", completePathToFile))
	cat("# R Syntax to import metadata to a data frame named theData. Run this syntax after defining theData by your data frame with the same variable names as in the original metadatafile (case does not matter).\n\n")
	for (i in returnObject){
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'label') <- '",i$varlabel,"'\n", sep="")
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'instruction') <- '",i$varinstruction,"'\n", sep="")
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'legalvalues') <- ",ifelse(toupper(i$varlegal != "ZEICHENKETTE")), paste("c(",i$varlegal, collapse=","),")\n", sep="")
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'missingvalues') <- c(",paste(i$varmissing, collapse=","),")\n", sep="")
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'legalvalueslabels') <- ",unlist(lapply(i$varlegallabels, function(x) x[2])[[1]]),"\n", sep="")
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'missingvalueslabels') <- list(", sep="")
			#",lapply(i$varmissinglabels, function(x) paste0("(",x[1],",",x[2],")")), "
		cat(")\n", sep="")
	}
 	sink()
}
