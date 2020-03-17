list.of.packages <- c("readxl", "rJava")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages, repos = "https://ftp.acc.umu.se/mirror/CRAN/")
library("readxl")

formats <- c("%m/%d/%y %I:%M:%OS %p", 
	     "%m/%d/%Y %H:%M", 
	     "%Y-%m-%d %H:%M:%OS", 
	     "%Y-%m-%d", 
	     "%d-%m-%y %H:%M:%OS",
	     "%d.%m.%y %H:%M:%OS",
	     "%d/%m/%y %H:%M",
	     "%d/%m/%y %I:%M:OS %p",
	     "%d/%m/%Y %H:%M",
	     "%d/%y/%m %H:%M %p",
	     "%d/%y/%m %H:%M:OS"
	     );

delist = function(y){
	return(vapply(y, paste, collapse = ', ', character(1L)));
}
specialCase <- function(entry, format){
	errEntry = NA;
	return(tryCatch({
		res = strptime(entry, format);
		return(res);
	}, warning = function(warning_condition){},
	error = function(error){
		print(error)
		return(NA);
	}))
}

alter <- function(entry){
	values <- lapply(formats, function(x){return(strptime(entry, x))})
	str <- substring(as.character(delist(entry)), 3)
	erEntry1 = specialCase(str, "%d-%m-%y %H:%M:%OS");
	if(is.na(erEntry1)) erEntry1 = specialCase(str, "%d-%m-%y");
	erEntry2 = specialCase(str, "%m-%d-%y %H:%M:%OS");
	if(is.na(erEntry2)) erEntry2 = specialCase(str, "%m-%d-%y");
	erEntry4 = specialCase(str, "%m-%d-%y %I:%M:%OS %p");
	arr <- strsplit(as.character(delist(entry)), " ")[[1]];
	str1 <- substring(arr[1], 0, nchar(arr[1])-4);
	str2 <- substring(arr[1], nchar(arr[1])-1, nchar(arr[1]))
	str1 <- paste(str1, str2, sep="");
	erEntry5 = specialCase(paste(str1, arr[2]), "%d-%y-%m %H:%M:%OS");
	
	values[length(values)+1] <- list(erEntry1);
	values[length(values)+1] <- list(erEntry2);
	values[length(values)+1] <- list(erEntry4);
	values[length(values)+1] <- list(erEntry5);
	values[length(values)+1] <- list(specialCase(substring(str, 4), "%y-%d-%m"));

	alt <- lapply(values, function(x){return(!is.na(x))})
	values = values[unlist(alt)];
	return(values)
}


recur <- function(entries, date){
	start <- date$start;
	end <- date$end;
	period <- difftime(date$start, date$end);
	corrected <- {};
	previous <- start;
	for(i in 1:length(entries)){
		opt = alter(entries[[i]]);
		opt = opt[sapply(opt, function(x){
					 isStart = abs(x$year - start$year) <= 1 
					 isEnd = abs(x$year - end$year) <= 1
					 return(isStart && (isEnd || is.na(isEnd)))
	     })]
		debug = opt;
		if(length(opt) > 1) {
			if(length(opt) > 1){
				opt <- unique(opt);
			}
			if(length(opt) > 1){
				alt <- sapply(opt, function(x){
						      return(abs(floor(difftime(x, delist(previous), units="day"))))

	     })			
				opt <- opt[alt == min(alt)];
			}
			if(length(opt) > 1){
				alt <- sapply(opt, function(x){
						      return(0 != (x$hour + x$min))
	     })
				if(length(opt[alt]) > 0){
					opt <- opt[alt];
				}
			}
		}
		if(length(opt) < 1){
			cat(paste("Row number ", i+4), "\n")
			print(date)
			print(entries[[i]])
			print(debug)
			print(opt)
			stop("ERROR missing Options")
		}
		if(is.na(opt)|| "NULL" %in% opt || "" == opt){
			cat(paste("Row number ", i+4))
			print(debug)
			print(opt)
			stop("ERROR is NA")
		}
		previous = opt;
		corrected$dates <- c(corrected$dates, opt)	
	}
	return(delist(corrected$dates))
}

fileDate <- function(fileName){
	tmp <- strsplit(fileName, " to ");
	start <- substring(tmp[[1]][1], 57)
	start <- strptime(start, "%Y%m%d");
	if(start$year < 0) start$year <- start$year + 2000;
	end <- substring(tmp[[1]][2], 1, 10)
	end <- strptime(end, "%Y%m%d");
	date <- {};
	date$end <- end;
	date$start <- start;
	return(date)			
}

outputFileName = NA;
noChange = 5;
directory = NA;
filePattern = NA;
files = NA;
while(is.na(directory)){
	allDirs <- list.dirs(path = ".", full.names = FALSE, recursive = FALSE);
	print(allDirs);
	cat(paste("Data location directory: "));
	con <- file("stdin");
	tmp <- readLines(con, n = 1L);
	if(!is.na(tmp)){
		directory <- allDirs[as.numeric(tmp)]
		cat(paste("Selected:", directory), "\n")
	} 
	close(con)
};	
while(length(files) == 0 || is.na(files)){
	files <- list.files(directory, full.names = TRUE);
	cat("No filter All files in Directory [y/n]: ")
	con <- file("stdin");
	res <- readLines(con, n = 1L);
	close(con)
	if(res %in% "n" || res %in% ""){
		cat("Files in directory [y/n]: ")
		con <- file("stdin");
		tmp <- readLines(con, n = 1L);
		if(tmp %in% "y") cat(files);
		close(con)

		cat(paste("File pattern [ Nuolja ] : "));
		con <- file("stdin");
		filePattern <- readLines(con, n = 1L);
		if(!is.na(filePattern)) filePattern <- "Nuolja"
		close(con)

		files <- files[grepl(filePattern, files)];

		if(length(files) == 0) files = NA;
	}
	# files <- files[grepl("Pole 18|Pole 17", files)];
	# files <- files[grepl("Pole 18", files)];
}
files <- lapply(files, function(file){
			res <- list(file, fileDate(file));
			return(res)
	     })
years <- sapply(files, function(file){
			return(file[[2]]$end$year +1900)
	     });
while(is.na(outputFileName)){
	cat("Year coverage: ", "\n")
	cat(unique(years), "\n")

	outputFileName <- directory;
	cat("Name output file [default:  directory name] : ");
	con <- file("stdin");
	outputFileName <- readLines(con, n = 1L);
	if(is.na(outputFileName) || outputFileName %in% "") outputFileName <- directory 
	close(con)
}



readFiles <- sapply(files, 
		    function(fileV){
			    file <- fileV[[1]]
			    date <- fileV[[2]]
			    cat(file, sep="\n");
			    mydf <- as.matrix(read_xlsx(file, cell_limits(c(4, 1), c(NA, 4)), sheet = 1, col_types = c("numeric", "list", "text", "text")));
			    mydf[,4] <- gsub("[^0-9.-]", "", delist(mydf[,4]))

			    dec <- as.matrix(read_xlsx(file, cell_limits(c(4, 5), 
									 c(nrow(mydf)+4, 5)), sheet = 1, col_types = c("list")));
			    dec <- gsub("[^0-9.-]", "", delist(dec))
			    if(length(dec) > 0){
				    for(i in 1:nrow(mydf)){
					    x = as.numeric(dec[i])
					    if(!is.na(x)){

						    mydf[i,4] = delist(as.numeric(delist(mydf[i,4])) + x/10)
					    }
				    }
			    }
			    mydf[,2] <- recur(mydf[,2], date);
			    return(mydf)
		    })

first <- TRUE;
combFile <- NA;
for(file in readFiles){
	if(first){
		first <- FALSE;
		combFile = data.frame(file);
		colnames(combFile) <- c("Pole", "Date/Time", "Unit", "Value");
	}else{
		colnames(file) <- c("Pole", "Date/Time", "Unit", "Value");
		combFile = rbind(combFile, file);
	}
}


combFile[,1] <- delist(combFile[,1]);
combFile[,2] <- delist(combFile[,2]);
combFile[,3] <- delist(combFile[,3]);
combFile[,4] <- delist(combFile[,4]);

date <- sapply(combFile[,2], function(x){
			   value = strsplit(x, " ")[[1]];
		     return(as.character(value[1]))
		    });
combFile <- cbind(combFile, date);
time <- sapply(combFile[,2], function(x){
			   value = strsplit(x, " ")[[1]];
		     return(as.character(value[2]))
		    });
combFile <- cbind(combFile, time);

write.csv(data.frame(combFile), paste(outputFileName, ".csv", sep=""));
cat("Wrote to file", sep="\n")
warnings()
