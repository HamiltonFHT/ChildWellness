#' BMI Analysis for custom search from PSS
#' Copyright (C) 2014  Tom Sitter - Hamilton Family Health Team
#' 
#' This program is free software; you can redistribute it and/or modify
#' it under the terms of the GNU General Public License as published by
#' the Free Software Foundation; either version 2 of the License, or
#' (at your option) any later version.
#' 
#' This program is distributed in the hope that it will be useful,
#' but WITHOUT ANY WARRANTY; without even the implied warranty of
#' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#' GNU General Public License for more details.


#' TODO:
#'     Confirm outlier criteria


#' Child Wellness Report Generators
#' Produces three registries:
#' BMI never measured
#' BMI 1 year out of date
#' BMI up-to-date (both height and weight measured within 1 year)
#' Produces three plots:
#'  # of patients in each of the above 3 registries
#'  # of up-to-date patients in each BMI category
#'  Boxplot for date of last weight and date of last height to see if they are being measured at same time.
#'  
#'  Outliers: < 3rd or > 99th percentile
#'            Increase or decrease by two percentiles between reports (?time between reports)

runReport <- function() {
  
  #+
  #Prompt user for file to read
  input_files = choose.files();
  
  if (length(input_files) == 0) {
    stop("No file selected");
  }

  #+ 
  #Prompt user for folder to save results to 
  output_dir = choose.dir(default=dirname(input_files[[1]]),caption="Select a directory to save files to");

  
  #If no output directory selected, using the working directory
  if (is.na(output_dir)) {
    output_dir = getwd();
    print("Saving to working directory: ");
    print(getwd());
  }
  
  age_range = getAgeRange()
  minAge = age_range[1]
  maxAge = age_range[2]

   
  #Comparison data between reports
  master_data = c()
  master_count = c()
  percentile_data = data.frame();
  
  num_files = length(input_files)
  
  if (num_files > 1) {
    pb <- winProgressBar(title="Progress...", min=0, max=num_files, width=300)
  }
  
  doctors_selected = FALSE;
  
  #Process each file
  for (i in 1:num_files) {
    
    if (num_files == 1) {
      current_file = input_files
    } else {
      current_file = input_files[i]
      if (exists("pb")) {
        setWinProgressBar(pb, i, title=sprintf("Progress (File %d of %d)", i, num_files))
      }
    }
    
    #Get filename
    filename = basename(sub("\\.txt","",current_file,fixed=FALSE))
    
        
    #Read Report File
    data = readReport(current_file)
    #Get the current date, and one year ago
    current_date = data$Current.Date[1]

    #If Age in years is not already calculated, do it here
    if ((!"Calc.Age" %in% colnames(data)) && ("Birth.Date" %in% colnames(data))){
      data$Calc.Age <- as.numeric((current_date - data$Birth.Date)/365.25)
    }
    
    if (!doctors_selected) {
      selected = getDoctors(unique(data$Doctor.Number));
      
      if (length(selected) > 1) {
        numbers = paste(selected, collapse = " ")
        doctors = sprintf("Doctors %s", numbers);
      } else if (length(selected) == 1) {
        doctors = sprintf("Doctor %s", selected);
      } else {
        stop("No Doctors Selected!");
      }
            
      doctors_selected = TRUE;
    }
    
    filename = paste(filename, doctors, sep="-");
    
    data = subset(data, data$Doctor.Number %in% selected);
        
    #Create Registries
    reg = getRegistries(data, minAge, maxAge, current_date)
    
    #If multiple files, store aggregate values in master data
    if (length(input_files) > 1) {
      total = nrow(reg$data)
      utd = sum(reg$up_to_date)
      master_data = rbind(master_data, c(current_date, total, utd, utd/total, sum(reg$never_done), sum(reg$out_of_date)))
      
      if (nrow(percentile_data) == 0) {
        percentile_data = reg$percentile;
      } else {
        #Get list of patient numbers for patients who moved two or more BMI groups between reports
        outlierPatients = findLargePercentileChanges(percentile_data, reg$percentile)
        df.temp = data.frame()
        #Look through each outlier patient
        for (p in outlierPatients) {
          #If patient already an outlier then add new reason to other reason(s)
          if (p %in% reg$outliers$Patient..) {
            index = (reg$outliers$Patient.. == p)
            reg$outliers[index,]$Reason <- 
                sprintf("%s, %s", reg$outliers[index,]$Reason, "Dramatic change in BMI")
          #If patient not already an outlier, then add them to a temporary registry
          } else {
            df.temp <- rbind(df.temp, reg$data[reg$data$Patient.. == p,])
          }
        }
        #If new outlier patients found, add them to the main outlier registry
        if (nrow(df.temp) > 0) {
          df.temp$Reason <-"Dramatic change in BMI"
          reg$outliers = rbind(reg$outliers, df.temp);
        }
      }
      
    }
    
    # Create Up-to-date over time chart

    
    #Print Graphs
    saveStatusGraph(output_dir, filename, minAge, maxAge,
                    sum(reg$never_done), sum(reg$up_to_date),sum(reg$out_of_date), nrow(reg$data),
                    current_date)
    
    bmi_count <- saveGrowthConcern(output_dir, filename, minAge, maxAge, reg, current_date)
    
    #Acummulate number of patients in each weight category
    master_count <- suppressWarnings(rbind(master_count, bmi_count))
    
    #Save height and weight data comparison chart(s)
    saveHeightWeightCharts(output_dir, filename, reg$data)

    #Print Registries to excel or csv depending on avilable libraries
    saveRegistries(output_dir, current_date, filename, reg)

  }
  
  #Save comparison data to file if multiple files read in
  if (length(input_files) > 1) {
    createMasterTable(output_dir, max(master_data[,1]),
                      master_data[order(master_data[,1]),],
                      master_count[order(master_count[,1]),],
                      minAge, maxAge,
                      doctors);
  }
  
  #Close and delete progress bar
  if (exists("pb")) {
    close(pb)
    rm(pb)
  }
  
  winDialog(type="ok",
            sprintf("Finished! You can find the files in %s", output_dir));
}
  
findLargePercentileChanges <- function(df1, df2) {
  
  #Row names are patient numbers
  #If both data frames have same patient number, check that they are within one percentile group of each other
  #or else they are outliers (no patient should change 2 BMI percentiles between reports)
  outliers = c()
  for (r in row.names(df1)) {
    #If patient is in both reports
    if (!is.na(df2[r,])) {
      #If they have moved more than one BMI group (0 == not measured, so exclude)
      if (abs(df2[r,] - df1[r,]) > 1 & df2[r,] != 0 & df1[r,] != 0) {
        #Add patient number to outlier list
        outliers = rbind(outliers, r)
      }
    }  
  }
  
  return(outliers)
}


getAgeRange <- function() {
  #get age ranges from user
#   minAge = as.numeric(winDialogString(message="What is the minimum age?", default="2"))
#   
#   if (minAge < 0) {
#    minAge = 2
#   } else if (minAge > 100) {
#    minAge = 100
#   } else if (is.na(minAge)) {
#    minAge = 2
#   }
#   
#   maxAge = as.numeric(winDialogString(message="What is the maximum age?", default="5"))
#   if (maxAge < 0) {
#     maxAge = 2
#   } else if (maxAge > 100) {
#     maxAge = 100
#   } else if (is.na(maxAge)) {
#     maxAge = 18
#   }
  
  age_range = menu(choices=c("2-5", "5-19"), graphics=TRUE, title="Select Age Range");
  if (age_range ==  0) {
    stop("No age range selected")
  }
    
  # 1: 2-5 years old    2: 5-19 years old 
  if (age_range == 1) {
    minAge = 2; maxAge=5;
  } else {
    minAge=5; maxAge=19;
  }
  
  return (c(minAge, maxAge)); 
}

getDoctors <- function(doctors) {
    
  selected = select.list(choices=as.character(doctors), 
                         preselect=as.character(doctors), 
                         multiple=TRUE, 
                         graphics=TRUE, 
                         title="Select Doctor Number(s)");  
  
  
  if (length(selected) == 0) {
    stop("No Doctors Selected!");
  }
  
  return(selected);
}

readReport <- function(input_file) {
    
    #Read in comma-separated text file
    data = read.csv(input_file)
    
    if (nrow(data) == 0) {
      winDialog(type="ok",
                sprintf("No data found in file %s!", input_file));
      stop("No data found in file")
    }
    
    # Delete auto-added "X" column
    if (!is.null(data$X)) {
      data <- data[,-(ncol(data))]
    }
    
    #+
    #' Convert data columns to R dates
    data$Date.of.Latest.Height = as.Date(data$Date.of.Latest.Height, format="%b %d, %Y")
    data$Date.of.Latest.Weight = as.Date(data$Date.of.Latest.Weight, format="%b %d, %Y")
    data$Date.of.Latest.BMI = as.Date(data$Date.of.Latest.BMI, format="%b %d, %Y")
    data$Date.of.Latest.BMI.Percentile = as.Date(data$Date.of.Latest.BMI.Percentile, format="%b %d, %Y")
    data$Current.Date = as.Date(data$Current.Date, format="%b %d, %Y")
    if ("Birth.Date" %in% colnames(data)) {
      data$Birth.Date = as.Date(data$Birth.Date, format="%b %d, %Y")
    }
    if ("Age" %in% colnames(data)) {
      data$Calc.Age = convertAge(data$Age)
    }
        
    # Convert BMI percentile to numberic values
    data$Latest.BMI.Percentile <- suppressWarnings(as.numeric(as.character(data$Latest.BMI.Percentile)))
    data$Latest.BMI <- suppressWarnings(as.numeric(as.character(data$Latest.BMI)))
  
    return(data)
}

#To convert ages to numbers
as.numeric.factor <- function(x) {as.numeric(levels(x))[x]}

#Convert age, accounting for "34mo" type ages
convertAge <- function(PSS_age) {
  
  if (is.numeric(PSS_age)) {
    return(PSS_age);
  }
  
  age_in_years = rep(0, length(PSS_age));
  
  for (i in 1:length(PSS_age)) {
    age = PSS_age[i]
    if (length(grep("mo", age)) > 0) { 
      age_in_years[i] = as.numeric(gsub("mo", "", age))/12;
    } else {
      age_in_years[i] = as.numeric.factor(age);
    }
  }
  
  return(age_in_years)
}

as.numeric.factor <- function(x) {suppressWarnings(as.numeric(levels(x))[x])}

getRegistries <- function(data, minAge, maxAge, current_date) {
    #+
    #' Create Registries
    #' May not add up to all patients due to outliers and data entry issues.
    
    #Subset based on user specified age range and privacy
    data = subset(data, data$Calc.Age >= minAge & data$Calc.Age < maxAge); # & data$Privacy != "Private Chart");
  
    if (nrow(data) == 0) {
      winDialog(type="ok",
                sprintf("No data found in age range %d to %d!\nRe-run with new age range?", minAge, maxAge));
      stop("No data found in age range")
    }
    
    one_year_ago = seq(current_date, length=2, by= "-12 months")[2]
    
    #' Outliers
    #' BMI Percentile < 3 or BMI Percentile > 99
    #' Date of measurement more recent than date of report
    
    #' Remove outliers from dataframe
    bottom_3rd_percentile = subset(data, data$Latest.BMI.Percentile < 3)
    if (nrow(bottom_3rd_percentile) > 0) {
      bottom_3rd_percentile$Reason <-"Bottom 3rd percentile"
    }
    
    top_99th_percentile = subset(data, data$Latest.BMI.Percentile > 99)
    if (nrow(top_99th_percentile) > 0) {
      top_99th_percentile$Reason <- "Top 99th percentile"
    }
    
    #' This is incorrect data which should be removed now
    incorrect_date = subset(data, data$Date.of.Latest.Height > current_date |
                               data$Date.of.Latest.Weight > current_date | 
                               data$Date.of.Latest.BMI > current_date)
    
    if (nrow(incorrect_date) > 0) {
      incorrect_date$Reason <- "Future Date?"
    }
    
    #Remove outliers from dataset
    data = data[!data$Patient.. %in% incorrect_date$Patient..,]
    
    num_patients = nrow(data)
    
    #Create TRUE/FALSE index for each registry
    out_of_date = never_done = out_of_date_never_done = up_to_date = rep(FALSE, num_patients);
    
    up_to_date[which(data$Date.of.Latest.Height > one_year_ago &
                       data$Date.of.Latest.Weight > one_year_ago)] = TRUE
    
    out_of_date[which((data$Date.of.Latest.Height <= one_year_ago | 
                             data$Date.of.Latest.Weight <= one_year_ago) &
                             !is.na(data$Date.of.Latest.BMI), arr.ind=TRUE)] = TRUE
    
    never_done[which(is.na(data$Date.of.Latest.BMI))] = TRUE
    
    out_of_date_never_done[which(out_of_date | never_done)] = TRUE
        
    #+ Get BMI status of up to date patients
    severely_wasted = wasted = normal = risk_of_overweight = overweight = obese = severely_obese = rep(FALSE, num_patients);

    #severely_wasted[which(up_to_date & data$Latest.BMI.Percentile<0.1)] = TRUE
    
    wasted[which(up_to_date & data$Latest.BMI.Percentile<=3)] = TRUE
    normal[which(up_to_date & data$Latest.BMI.Percentile>3 & data$Latest.BMI.Percentile<=85)] = TRUE
    
    if (maxAge == 5) {
      risk_of_overweight[which(up_to_date & data$Latest.BMI.Percentile>85 & data$Latest.BMI.Percentile<=97)] = TRUE
      overweight[which(up_to_date & data$Latest.BMI.Percentile>97 & data$Latest.BMI.Percentile<=99.9)] = TRUE
      obese[which(up_to_date & data$Latest.BMI.Percentile>99.9)] = TRUE
      
      percentile = rep(0, num_patients);
      percentile[wasted] = 1
      percentile[normal] = 2
      percentile[risk_of_overweight] = 3
      percentile[overweight] = 4
      percentile[obese] = 5
    } else {
      overweight[which(up_to_date & data$Latest.BMI.Percentile>85 & data$Latest.BMI.Percentile<=97)] = TRUE
      obese[which(up_to_date & data$Latest.BMI.Percentile>97)] = TRUE
      severely_obese[which(up_to_date & data$Latest.BMI.Percentile>99.9)] = TRUE
      
      percentile = rep(0, num_patients);
      percentile[wasted] = 1
      percentile[normal] = 2
      percentile[overweight] = 3
      percentile[obese] = 4
    }
    
    at_risk = (up_to_date & !normal)
    
    percentile_df = data.frame(percentile, row.names=data$Patient..);
    
    registries <- list("data" = data,
                       "outliers" = rbind(incorrect_date, bottom_3rd_percentile, top_99th_percentile),
                       "percentile" = percentile_df,
                       "up_to_date" = up_to_date,
                       "out_of_date" = out_of_date,
                       "never_done" = never_done,
                       "out_of_date_never_done" = out_of_date_never_done,
                       #"severely_wasted" = sum(severely_wasted),
                       "wasted" = sum(wasted),
                       "normal" = sum(normal),
                       "risk_of_overweight" = sum(risk_of_overweight),
                       "overweight" = sum(overweight),
                       "obese" = sum(obese),
                       "severely_obese" = sum(severely_obese),
                       "at_risk" = at_risk)

    return(registries)
}


saveStatusGraph <- function(output_dir, filename, minAge, maxAge,
                            num_never_done, num_up_to_date, num_out_of_date, num_total, 
                            current_date) {
  
  # Store counts of number of patients in each registry in array
  status_counts = c(num_total, num_up_to_date, num_out_of_date, num_never_done)
  status_labels = c("Total", "Up to Date", "Out of Date", "Never Done")
  status_colours = c("mediumpurple2", "darkolivegreen3", "orangered3", "dodgerblue3")
  
  full_filename = sprintf("%s/BMI_Status-%s.png", output_dir, filename)
  
  png(full_filename)
  # Make left margin larger for legend text
  par(mar = c(5,8,4,2) + 0.1);
  # Horizontal bar chart
  bp_status <- barplot(status_counts, col=status_colours, horiz=TRUE,
                       legend.text=status_labels,
                       xlab="Number of patients",
                       main=sprintf("Total Peds %d to %d years (n=%d)\nas of %s", 
                                    minAge, maxAge, num_total, format(current_date, "%b %d, %Y")),
                       axes=F, xlim=c(0,ceiling(max(status_counts)/100)*100));

  # Make an adjusted vector for bar count positioning
  adjusted_count = c()
  index = 1
  for(i in status_counts) {if (i < 25) {adjusted_count[index]=i*2+40} 
                        else {adjusted_count[index]=i}
                        index = index+1
  }
  
  # Add axis labels
  axis(1, at = seq(0,ceiling(max(status_counts)/100)*100,100));
  axis(2, at = bp_status, labels=status_labels, las=1);
  text(x=adjusted_count/2, y=bp_status,
       labels=as.character(status_counts), xpd=TRUE)
  dev.off();
}

saveGrowthConcern <- function(output_dir, filename, minAge, maxAge, reg, current_date) {

  # Store results in vectors
  if (maxAge == 5) {
    bmi_labels = c("Wasted", "Normal", "Risk of Overweight", "Overweight", "Obese");
    bmi_counts <- c(#reg$severely_wasted, 
      reg$wasted, reg$normal, reg$risk_of_overweight, reg$overweight, reg$obese);
  } else {
    bmi_labels = c("Wasted", "Normal", "Overweight", "Obese", "Severely Obese");
    bmi_counts <- c(#reg$severely_wasted, 
      reg$wasted, reg$normal, reg$overweight, reg$obese, reg$severely_obese);
  }
  
  # "dodgerblue3"
  bmi_colours = c("orangered3", "darkolivegreen3", "mediumpurple2", "mediumturquoise", "orange", "lightskyblue");
  
  full_filename = sprintf("%s/BMI Count-%s.png", output_dir, filename);
  png(full_filename);
  
  # Make left margin larger for legend text
  par(mar = c(5,4,6,2) + 0.1);
  # Plot BMI status
  bp_bmi <- barplot(bmi_counts, 
                    main=sprintf("Total Patients %d to %d years with up to data BMI (n=%d/%d)\nas of %s",
                                 minAge, maxAge, sum(reg$up_to_date), nrow(reg$data), format(current_date, "%b %d, %Y")),
                    ylab="Number of patients", 
                    col=bmi_colours,
                    legend.text=bmi_labels,
                    las=2,
                    axes=F, ylim=c(0,ceiling(max(bmi_counts)/50)*50));
  text(bp_bmi, 
       par("usr")[3], 
       labels=bmi_labels, 
       srt=40, xpd=TRUE, adj=c(1.1, 1.1), cex=0.9);
  
  # Plot y axis
  axis(2, at=seq(0,ceiling(max(bmi_counts)/50)*50,50))
  
  # Make an adjusted vector for bar count positioning
  adjusted_count = c()
  index = 1
  for(i in bmi_counts) {if (i < 10) {adjusted_count[index]=i*2+10} 
                        else {adjusted_count[index]=i}
                        index = index+1
  }
  
  
  text(y=adjusted_count/2, x=bp_bmi, 
       labels=as.character(bmi_counts), xpd=TRUE)
  dev.off();
  return(c(as.numeric(current_date), bmi_counts))
}

saveHeightWeightCharts <- function(output_dir, filename, data) {
  
  #Plot Date of Latest Height vs Date of Latest Weight
  full_filename = sprintf("%s/HeightWeightBoxplot-%s.png", output_dir, filename)
  png(full_filename);
  
  boxplot(data$Date.of.Latest.Height, data$Date.of.Latest.Weight, 
          names=c("Height", "Weight"),
          col=c("darkolivegreen3", "lightskyblue"), 
          main="Date of Latest Height and Weight")
  
  dev.off()
  
  # Previously used scatter plot
  
  # HvW_data = data[!data$Patient.. %in% outliers$Patient..,]
  # png(filename=paste(output_dir, "HeightvsWeight.png", sep="/"));
  # plot(HvW_data$Date.of.Latest.Weight, HvW_data$Date.of.Latest.Height, 
  #      xaxt="n", yaxt="n",
  #      main="Date of Latest Weight vs. Height",
  #      xlab="Date of Latest Weight", 
  #      ylab="Date of Latest Height")
  # axis.Date(side = 2, x=HvW_data$Date.of.Latest.Height, format = "%Y")
  # axis.Date(side = 1, x=HvW_data$Date.of.Latest.Weight, format = "%Y")
  # abline(a=0, b=1, col="green")
  # dev.off()
  
}

saveRegistries <- function(output_dir, current_date, filename, reg) {
      
    #' Prepare to save registries. Check to make sure xlsx library is installed and install if necessary
    #' Write to a CSV text file otherwise.
    if ("xlsx" %in% rownames(installed.packages())) {
      require(xlsx)
      writeToExcel(output_dir, filename, 
                   reg$data[reg$out_of_date_never_done,], 
                   reg$data[reg$at_risk,], 
                   reg$outliers)
    } else {
      
      response = winDialog(type="yesno", 
                           "Did not find Excel libraries. Would you like to install them now?\nYou must be connected to the internet and have Excel installed on this computer")
      
      if (response == "YES") {
        
        install.packages("xlsx")
        require(xlsx)
        
        if ("xlsx" %in% rownames(installed.packages())) {
          writeToExcel(output_dir, current_date, filename, 
                       reg$data[reg$out_of_date_never_done,], 
                       reg$data[reg$at_risk,], 
                       reg$outliers)
        } else {
          winDialog(type="ok", 
                    "Something went wrong installing excel libraries ('xlsx'). Writing to text files.");
          writeToCSV(output_dir, current_date, filename, 
                     reg$data[reg$out_of_date_never_done,], 
                     reg$data[reg$at_risk,], 
                     reg$outliers)
        }
      }
      else {
        writeToCSV(output_dir, current_date, filename, 
                   reg$data[reg$out_of_date_never_done,], 
                   reg$data[reg$at_risk,], 
                   reg$outliers)
      }
  }
}


createMasterTable <- function(output_dir, lastDate, master_data, master_count, minAge, maxAge, doctors) {
  
  # Turn data into a data frame
  df.MD <- data.frame(master_data)
  df.MC <- suppressWarnings(data.frame(master_count))
  
  # Turn lastDate into a Date
  lastDate <- as.Date(lastDate, origin="1970-01-01")
  
  # Give column names
  age_string = sprintf("Total Patients %syrs to %syrs", minAge, maxAge)
  colnames(df.MD) <- c("Date of Data Capture", age_string, 
                    sprintf("%s up to date BMI (12 months)", age_string), "Percentage",
                    sprintf("%s w/o BMI (never done)", age_string),
                    sprintf("%s w/ out-of-date BMI", age_string))
  df.MD$"Date of Data Capture" <- as.Date(df.MD$"Date of Data Capture", origin="1970-01-01")
  
  if (maxAge == 5) {
    colnames(df.MC) <- c("Date of Data Capture", #"Severely Wasted", 
                         "Wasted", "Normal",
                         "Risk of Overweight", "Overweight", "Obese")
  } else {
    colnames(df.MC) <- c("Date of Data Capture", #"Severely Wasted", 
                         "Wasted", "Normal",
                          "Overweight", "Obese", "Severely Obese")
  }
  
  
  df.MC$"Date of Data Capture" <- as.Date(df.MC$"Date of Data Capture", origin="1970-01-01")
  
  numColsMC = length(df.MC);
  numColsMD = length(df.MD);
  
  # Create workbook and title
  excel_file <- sprintf("%s/Child Wellness Master Table %s %s.xlsx", output_dir, lastDate, doctors)
  outwb <- createWorkbook(type="xlsx")
  
  # Create Summary Sheet
  sheet.MD <- createSheet(outwb, sheetName="Summary")
  setColumnWidth(sheet.MD, 1:numColsMD, 15)
  
  csPerc <- CellStyle(outwb, dataFormat=DataFormat("0.00%"))
  csWrap <- CellStyle(outwb, alignment=Alignment(wrapText=T))

  df.MD.colPerc <- list('4'=csPerc)
  addDataFrame(df.MD, sheet.MD, colStyle=c(df.MD.colPerc), row.names=F)
  
  # Word wrap the header rows
  row <- getRows(sheet.MD, rowIndex=1)
  cell <- getCells(row)
  for (i in 1:numColsMD){
    setCellStyle(cell[[paste('1.',i, sep="")]], csWrap)
  }
  
  # Create count sheet
  sheet.MC <- createSheet(outwb, sheetName="Count")
  setColumnWidth(sheet.MC, 1:numColsMC, 12)
  csCenter <- CellStyle(outwb, alignment=Alignment(h="ALIGN_CENTER"))
  df.MC.colCenter <- list('2'=csCenter, '3'=csCenter, '4'=csCenter,
                          '5'=csCenter, '6'=csCenter, '7'=csCenter)
  addDataFrame(df.MC, sheet.MC, colStyle=c(df.MC.colCenter), row.names=F)
  row <- getRows(sheet.MC, rowIndex=1)
  cell <- getCells(row)
  for (c in 1:numColsMC){
    setCellStyle(cell[[sprintf('1.%d', c)]], csWrap)
  }
  
  # Create Percent Change sheet
  sheet.PC <- createSheet(outwb, sheetName="Percent Change")
  df.PC.colPerc <- list('2'=csPerc, '3'=csPerc, '4'=csPerc,'5'=csPerc)
  df.PC <- df.MD[,-4]
  numColsPC = length(df.PC)
  setColumnWidth(sheet.PC, 1:numColsPC, 15)  
  df.PC.final <- c()
  for (i in 1:(length(df.PC[[1]])-1)) {
    df.PC.final <- rbind(df.PC.final, cbind(df.PC[[1]][i+1],(df.PC[i+1,2:numColsPC]-df.PC[i,2:numColsPC])/df.PC[i,2:numColsPC]))
  }
  colnames(df.PC.final)[1]<-"Date of Data Capture" 
  addDataFrame(df.PC.final, sheet.PC, colStyle=c(df.PC.colPerc), row.names=F)
  row <- getRows(sheet.PC, rowIndex=1)
  cell <- getCells(row)
  for (i in 1:numColsPC){
    setCellStyle(cell[[sprintf('1.%d',i)]], csWrap)
  }
  
  # Save Workbook
  saveWorkbook(outwb, excel_file)
  
  filename <- sprintf("Up-To-Date Over Time Chart (%syrs to %syrs) %s %s.png", minAge, maxAge, lastDate, doctors)
  xrange <- range(df.MD$"Date of Data Capture")
  yrange <- range(df.MD$"Percentage")
  png(filename=paste(output_dir,filename,sep="/"), width=700, height=480)
  par(mar=c(5.5, 5.5, 4.1, 2.1), mgp=c(4, 1, 0))
  heading <- sprintf("Percent of Up-to-Date Over Time (%s years to %s years)", minAge, maxAge)
  plot(df.MD$"Date of Data Capture", df.MD$"Percentage", main=heading, ylab=expression(bold(Percent~of~patients~with~up-to-date~BMI)),
       xlab=expression(bold(Date)), pch=23, col="blue", bg="blue", yaxt="n", ylim=c(round(yrange[1],2)-0.01,round(yrange[2],2)+0.01),
       xlim=c(xrange[1]-15,xrange[2]+15), xaxt="n", font=2)
  text(df.MD$"Date of Data Capture", df.MD$"Percentage", sprintf("%.2f%%", df.MD$"Percentage"*100), pos=1, col="blue")
  lines(df.MD$"Date of Data Capture", df.MD$"Percentage", type="o", col="blue", lwd=1.5)
  axis(2,at=seq(round(yrange[1],2)-0.01, round(yrange[2],2)+0.01, 0.005),
       labels=sprintf("%.2f%%", seq(round(yrange[1],2)-0.01, round(yrange[2],2)+0.01, 0.005)*100),
       las=2, cex.axis=0.85)
  axis.Date(1, df.MD$"Date of Data Capture", at=seq(xrange[1]-15,xrange[2]+15,by="months"), "%B-%d-%Y")
  
  dev.off()
}

#Write registries to three seperate CSV files
writeToCSV <- function(output_dir=getwd(), current_date=Sys.Date(), filename, out_of_date_never_done=data.frame(), at_risk=data.frame(), outliers=data.frame()) {
  write.csv(out_of_date_never_done,
            file=sprintf("%s/CW-Out Of Date-%s %s.txt", output_dir, filename, format(current_date, "%d%b%Y")),
            row.names=FALSE)
  write.csv(at_risk[order(at_risk$Latest.BMI.Percentile, decrease=TRUE), ],
            file=sprintf("%s/CW-At Risk-%s %s.txt", output_dir, filename, format(current_date, "%d%b%Y")),
            row.names=FALSE)
  write.csv(outliers,
            file=sprintf("%s/CW-Outliers-%s %s.txt", output_dir, filename, format(current_date, "%d%b%Y")),
            row.names=FALSE)
}

#Write registries to excel file (requires R xlsx package)
writeToExcel <- function(output_dir=getwd(), filename, out_of_date_never_done=data.frame(), at_risk=data.frame(), outliers=data.frame()) {
  excel_file = sprintf("%s/Child Wellness Registries %s.xlsx", output_dir, filename)
  
  if (nrow(out_of_date_never_done) != 0) {
    write.xlsx(out_of_date_never_done,
               file=excel_file, 
               sheetName="Out Of Date", 
               row.names=FALSE);
  }
  if (nrow(at_risk) != 0) {
    write.xlsx(at_risk, 
               file=excel_file, 
               sheetName="At Risk",
               row.names=FALSE,
               append=TRUE);
  }
  if (nrow(outliers) != 0) {
    write.xlsx(outliers, 
               file=excel_file, 
               sheetName="Outliers",
               row.names=FALSE,
               append=TRUE);
  }

}
#+
#' Run Application using this command
runReport()
