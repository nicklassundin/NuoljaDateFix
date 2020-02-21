list.of.packages <- c("readxl", "rJava")
# list.of.packages <- c("xlsx", "rJava")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages, repos = "https://ftp.acc.umu.se/mirror/CRAN/")
library("readxl")
# library("xlsx")
files <- list.files("IButtoon Data", full.names = TRUE);
files <- files[grepl("Nuolja", files)];
# files <- files[grepl(" 4 2018.09.28 to 2019.09.27", files)];
# files <- files[grepl("36|37|38|39", files)];
# files <- files[grepl("45 2018.09.28 to 2019.09.27.xlsx", files)];
# files <- files[grepl("57", files)];
# files <- files[grepl("63", files)];
# files <- files[grepl("36", files)];
# files <- files[grepl("32", files)];
# files <- files[grepl("32|36|37|57|63|71", files)];
# files <- files[grepl("37", files)];
# files <- files[grepl("71", files)];
# files <- files[grepl("Pole 7 2018.09.28 to 2019", files)];
# files

formats <- c("%m/%d/%y %I:%M:%OS %p", 
	     "%m/%d/%Y %H:%M", 
	     "%Y-%m-%d %H:%M:%OS", 
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
	# print(entry)
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
	str <- paste(str1, arr[2])
	erEntry5 = specialCase(str, "%d-%y-%m %H:%M:%OS");
	values[length(values)+1] <- list(erEntry1);
	values[length(values)+1] <- list(erEntry2);
	values[length(values)+1] <- list(erEntry4);
	values[length(values)+1] <- list(erEntry5);
	alt <- lapply(values, function(x){return(!is.na(x))})
	values = values[unlist(alt)];
	return(values)
}


recur <- function(entries, date){
	start <- date$start;
	# start <- strptime(date$start, "%Y.%m.%d");
	end <- date$end;
	# end <- strptime(date$end, "%Y.%m.%d");
	period <- difftime(date$start, date$end);
	corrected <- {};
	previous <- start;
	for(i in 1:length(entries)){
		opt = alter(entries[[i]]);
		opt = opt[sapply(opt, function(x){
					 return(x$year == start$year || x$year == end$year)
	     })]
		if(length(opt) > 1) {
			alt <- sapply(opt, function(x){
					      bool <- FALSE;
					      if(i > 1){
						      startDif = difftime(x, delist(previous), units="day")
						      bool <- startDif > 0
					      }else{
						      start = abs(difftime(x, date$start, units="day")) 
						      end = abs(difftime(x, date$end, units="day"))
						      bool <- abs(start + end) - abs(period) == 0;  
					      }
					      return(bool)
	     })

			opt = opt[alt];
			if(length(opt) > 1){
				opt <- unique(opt);
			}
			if(length(opt) > 1){
				alt <- sapply(opt, function(x){
						      return(difftime(x, delist(previous), units="day"))

	     })			
				opt <- opt[alt == min(alt)];
			}
		}
		if(length(opt) < 1){
			cat(paste("Row number ", i+4))
			stop("ERROR missing Options")
		}
		if(is.na(opt)){
			cat(paste("Row number ", i+4))
			stop("ERROR is NA")
		}
		previous = opt;
		corrected$dates <- c(corrected$dates, opt)	
	}
	return(delist(corrected$dates))
}

fileDate <- function(fileName){
	tmp <- strsplit(fileName, " to ");
	start <- substring(tmp[[1]][1], 38)
	start <- strptime(start, "%Y.%m.%d");
	if(start$year < 0) start$year <- start$year + 2000;
	end <- substring(tmp[[1]][2], 1, 10)
	end <- strptime(end, "%Y.%m.%d");
	date <- {};
	date$end <- end;
	date$start <- start;
	return(date)			
}

files <- lapply(files, function(file){
			res <- list(file, fileDate(file));
			return(res)})

years <- sapply(files, function(file){
			return(file[[2]]$end$year +1900)
			});
year <- NA;
while(!(year %in% years)){
	cat("available years", sep="\n");
	if(length(unique(delist(years))) > 1){
		cat(unique(delist(years)), sep="\n")
		cat("choose year: ");
		year <- as.numeric(readLines(file("stdin"), n = 1L));
	}else{
		year <- unique(delist(years))[1];
		cat("Only one option", sep="\n");
	}
}

files <- files[years %in% year]

readFiles <- sapply(files, 
		    function(fileV){
			    file <- fileV[[1]]
			    date <- fileV[[2]]
			    cat(file, sep="\n");
			    mydf <- as.matrix(read_xlsx(file, cell_limits(c(4, 1), c(NA, 5)), sheet = 1, col_types = c("numeric", "list", "text", "numeric", "numeric")));
			    dec <- delist(mydf[,5]);
			    if(any(!is.na(dec))){
				    for(i in 0:nrow(mydf)){
					    value = as.numeric(dec[i])
					    if(value < 10 && length(value) && !is.na(value) != 0) mydf[i,4] <- as.numeric(mydf[i,4]) + (value/10)
				    }
			    }
			    mydf <- mydf[,-5];
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

write.csv(data.frame(combFile), paste("Nuolja_", paste(year, ".csv", sep=""), sep=""));
cat("Wrote to file", sep="\n")
