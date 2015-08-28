# save excel sheet as "unicode text" or "Text (MS-DOS)" with extension *.txt
# enter directory and filename below and then source this script or highlight everything and run.
# the result will be a file with the name <filename.kodebuch> in the original folder.

# dir: directory in which the excel sheet resides  e.g., dir="C:/some path to/"
# filename:  #name of the txt file e.g., filename="sometxtfileextractedfromexcel.txt"

# results will be in a newly created folder named converted_<filename without *txt extension> in the indicated directory

version <- "\n\n..............................................\nfunction makeKodebuch, last change 20150827\nj.jacoby@iwm-tuebingen.de\n..............................................
makeKodebuch(dir=\".\", filename=NULL, showLog=TRUE)

> dir is the directory in which the text file resides, default: getwd()
> filename is the name of the file, required

the file must be a tab separated textfile made with xls2kodebuch.xlsm
			[available at http://j.mp/iwmfds]
\n
> THIS IS AN EXPERIMENTAL VERSION OF THIS SCRIPT. NO GUARANTEE IS BE GIVEN,
PLEASE CHECK OUTPUT FILES YOURSELF BEFORE COMMITTING THEM."
cat(version)


makeKodebuch <- function(dir=".", filename=NULL, showLog=TRUE) { 
	pureDir <- gsub("/+$","", dir)
	completePathToFile <- paste0(pureDir,"/",filename)
	#print(completePathToFile)
	
	if (!file.exists(completePathToFile)) { stop (paste0("The indicated file: ",filename," does not exist in ",dir,".\nPlease check whether path and file name are correct.")) }
  
	rangeFormat <- "^[0-9\\.\\-]+\\-[0-9\\.\\-]+$"
	valueFormat <- "^\\-?[0-9\\.]+\\;\\-?[0-9\\.]+"
	
	################################################
	origtext <- read.table(completePathToFile, quote="", sep="\t", stringsAsFactors=FALSE)
	names(origtext) <- c("name", "label", "instruction", "legal", "missing", "legallabels", "missinglabels", "spsstype")
  returnObject <- list()
  for (i in 1:nrow(origtext)){
  	#cat("\n",i)
  	parts <- origtext[i,]
  	
  	varlabelC <- gsub("^[[:blank:]]*", "", gsub("[[:blank:]]*$", "", parts$label))
  	varlegalC <- parts$legal
  		if (toupper(parts$legal)=="ZEICHENKETTE") { varlegalC <- "Zeichenkette" }
  		else if (length(grep(valueFormat, parts$legal)) > 0 && grep(";", parts$legal) == 1) { varlegalC <- as.numeric(strsplit(parts$legal, split=";")[[1]]) }

  	varmissingC <- as.numeric(strsplit(as.character(parts$missing), split=";")[[1]])
  	varlegallabelsC <- lapply(strsplit(parts$legallabels, split=";")[[1]], function(x) strsplit(x,split="=")[[1]])
  	varmissinglabelsC <- lapply(strsplit(parts$missinglabels, split=";")[[1]], function(x) strsplit(x,split="=")[[1]])

  	varlevelC <- ifelse(parts$spsstype=="--", NA, parts$spsstype)
  	valuetypeC <- ifelse(
	    		toupper(parts$legal) == "ZEICHENKETTE" , "STRING"
	    		, ifelse (length(grep(rangeFormat, as.character(parts$legal))) > 0, "RANGE"
	    				, "VALUES")
  		)

  	legalvaluesexpandC <- ifelse(
	  	length(varlegallabelsC) == 1
	    , paste0(varlegalC,' "', varlegallabelsC,'"')
	    , paste(sort(unlist(lapply(varlegallabelsC, function(x) paste0(x[1], ' "',x[2],'"')))), collapse="\n")
	  )
  	if (valuetypeC == "STRING") { legalvaluesexpandC <- paste0("Zeichenkette \"", parts$legallabels,"\"") }
 	
 	  missingvaluesexpandCv <- c()
  	
	  for (j in varmissinglabelsC) { missingvaluesexpandCv <- c(missingvaluesexpandCv, paste0(j[1],' "',j[2],'"')) }
	  missingvaluesexpandC <- paste(missingvaluesexpandCv, collapse="\n")
  	
#   	print(varmissinglabelsC)
#   	cat("\n----------------------\n")
#   	print(missingvaluesexpandC)
  	
  	
    newVar <- list(
      varname=toupper(parts$name)
      , varlabel=varlabelC
      , varinstruction = parts$instruction
      , varlegal=varlegalC
      , varmissing=varmissingC
      , varlegallabels=varlegallabelsC
      , varmissinglabels=varmissinglabelsC
    	, varlevel=varlevelC
    	, valuetype=valuetypeC
    	, legalvaluesexpand=legalvaluesexpandC
    	, missingvaluesexpand=missingvaluesexpandC
    )
    returnObject[[parts$name]] <- newVar
  }
	
  kodebuchfile <- gsub("\\..+",".kodebuch", paste0("/",filename))
  kodebuchfilestringarray <- c()
  for (i in returnObject){
  	#print(i)
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
  
  kodebuchfilestringarray <- c(kodebuchfilestringarray, 
                               paste(
                                 i$varname
                                 , paste0(
                                 		i$varlabel
                                 		, ifelse(!is.na(i$varlevel) && !is.null(i$varlevel)
                                 			, paste0(" [",i$varlevel,"]")
                                 			, "")
                                 	)
                                 , paste0("\"",i$varinstruction,"\"")
                                 , ifelse(
                                 		i$varlegal[1] == "Zeichenkette"
                                 		, "{Zeichenkette}"
                                 		, paste0("{",paste(i$varlegal, collapse=";"),"}")
                                 	)
                                 , paste0("{",paste(i$varmissing, collapse=";"),"}") 
                                 , i$legalvaluesexpand
                                 , i$missingvaluesexpand
                               	 , i$missingvaluesexpandC
                                 , sep="\n"
                               )
  														)
  }
	
  dirpath <- paste0("converted_",gsub(".txt$","", filename)) 
 	if (!file.exists(paste0(pureDir,"/",dirpath))) { dir.create(paste0(pureDir,"/",dirpath)) }
	cat(paste(kodebuchfilestringarray, collapse="\n\n"), file=paste0(pureDir,"/",dirpath,"/",kodebuchfile), append=FALSE)
	cat("directory:\n\t",paste0(pureDir,"/",dirpath),"\n")
	cat("files created:\n")
	cat(paste0("\t",gsub("^\\/","", kodebuchfile), "\n"))
	
	#################################################
	### create spss syntax per variable
	sink(file=paste0(pureDir,"/",dirpath,"/",gsub("\\.[[:alnum:]]+$", "_perVariable.sps", filename)), append=FALSE)
	
	cat("* SPSS Syntax to import metadata to an SPSS data file in per variable order. Run this syntax on the data set.\n\n")
	for (r in returnObject){ 
		cat ("VARIABLE LABELS ", r$varname, " '", r$varlabel,"/", r$varinstruction,"'.\n", sep="")
  		if (!is.na(r$valuetype) && r$valuetype == "VALUES")  { 
  			cat ("VALUE LABELS ",r$varname, " ", gsub("[[:space:]]*\"", "'", gsub("\n"," ", r$legalvaluesexpand)), ".\n", sep="")
  		}
		cat ("MISSING VALUES  ", r$varname, " (", paste(r$varmissing, collapse=","), ").\n", sep="")
		
		if (!is.null(r$varlevel) && !is.na(r$varlevel)) { cat ("VARIABLE LEVEL  ",r$varname, " (", r$varlevel, ").\n", sep="") }
		
		cat ("VARIABLE ATTRIBUTE  VARIABLES=",r$varname," ATTRIBUTE=MissingDefinition ('",paste(unlist(lapply( r$varmissinglabels, function(x) { paste0(x[1],"=", x[2])} )), collapse=";"),"').\n", sep="")
		cat(ifelse(r$valuetype=="STRING","STRING ","NUMERIC "), r$varname,".\n", sep="")
		cat("\n")
	}
	sink()
	cat(paste0("\t",gsub("\\.[[:alnum:]]+$", "_perVariable.sps", filename), "\n"))

	
	
	####################################################
	### create spss syntax per type
	sink(file=paste0(pureDir,"/",dirpath,"/",gsub("\\.[[:alnum:]]+$", "_perType.sps", filename)), append=FALSE)
	cat("* SPSS Syntax to import metadata to an SPSS data file in per type order. Run this syntax on the data set.\n\n")
	
	cat("VARIABLE LABELS\n")
	for (i in returnObject){cat (" ",i$varname, " '",i$varlabel,"/",i$varinstruction,"'\n", sep="")}
	
	cat(".\nVALUE LABELS\n")
	counter <- 1
	for (i in returnObject)	{ 
		if ( i$valuetype == "VALUES")  { 
			cat (ifelse(counter==1," "," /"),i$varname, " ", gsub("[[:space:]]*\"", "'", gsub("\n"," ", i$legalvaluesexpand)), "\n", sep="") 
			counter <- counter + 1
		}
	}
	
	cat(".\nMISSING VALUES\n")
	for (i in returnObject) { cat (" ",i$varname, " (", paste(i$varmissing, collapse=","), ")\n", sep="") }
	
	cat(".\n")
	if (!is.null(i$varlevel) && !is.na(i$varlevel)) { 
		cat ("VARIABLE LEVEL\n")
		for (i in returnObject) {cat (" ",i$varname, " (", i$varlevel, ")\n", sep="")}
	}
	cat(".\n\n")
	
	counter <- 0
	cat("VARIABLE ATTRIBUTE\n")
	for (i in returnObject) {
			counter <- counter+1
			cat (ifelse(counter==1, "", "/"),"VARIABLES=",i$varname," ATTRIBUTE=MissingDefinition ('",paste(unlist(lapply( i$varmissinglabels, function(x) { paste0(x[1],"=", x[2])} )), collapse=";"),"')\n", sep="")
	}
	cat(".\n\n")
	for (i in returnObject) {	cat(ifelse(i$valuetype=="STRING","STRING ","NUMERIC "),i$varname,".\n", sep="")}
	
	sink()
	cat(paste0("\t",gsub("\\.[[:alnum:]]+$", "_perType.sps", filename), "\n"))
	
	
	########################################
	### create R syntax
	sink(file=paste0(pureDir,"/",dirpath,"/",gsub("\\.[[:alnum:]]+$", "_forR.R", filename)), append=FALSE)
	cat("# R Syntax to import metadata to a data frame named theData. Run this syntax after defining theData by your data frame with the same variable names as in the original metadatafile (case does not matter).\n\n")
	for (i in returnObject){
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'label') <- '",i$varlabel,"'\n", sep="")
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'instruction') <- '",i$varinstruction,"'\n", sep="")
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'legalvalues') <- ",
			ifelse(
				i$valuetype == "VALUES"
				, paste0("c(",paste(i$varlegal, collapse=","),")\n")
				, paste0("'",i$varlegal,"'\n")
			)
			, sep=""
		)
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'missingvalues') <- c(",paste(i$varmissing, collapse=","),")\n", sep="")
		
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'legalvalueslabels') <- ",
			ifelse( i$valuetype=="VALUES"
				, gsub("\"","'", gsub("\\(\\\"([\\-\\.0-9]+)\\\",", "(\\1,", paste0("list(", paste0(i$varlegallabels, collapse=", "),")")))
				, paste0("c('", i$varlegallabels,"')")
			)
			,"\n"
			, sep=""
		)
		cat ("attr(theData[,names(theData) == '",i$varname,"'], 'missingvalueslabels') <- "
			, ifelse(length(i$varmissinglabels) > 1
				, gsub("\"","'", gsub("\\(\\\"([0-9\\-]+)\\\",", "(\\1,", paste0("list(", paste0(i$varmissinglabels, collapse=","),")")))
				#, gsub("\"","'", paste0("list(",i$varmissinglabels,")"))
				, paste0("list(",gsub("\\\")", "')", gsub("\\\", \\\"", ", \'", gsub("\\(\"","(", i$varmissinglabels))),")")
			)
			, "\n"
			,sep=""
		)
		for (j in i$varmissinglabels) { cat ("theData[,'",i$varname,"'][theData[,'",i$varname,"'] == ",j[1],"] <- NA\n",sep="") }
	}
 	sink()
	cat(paste0("\t",gsub("\\.[[:alnum:]]+$", "_forR.R", filename), "\n"))
	#if (sink.number() > 0) {for (i in 1:sink.number()) {sink()} }
}
