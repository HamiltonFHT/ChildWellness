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
#' 
#' You should have received a copy of the GNU General Public License along
#' with this program; if not, write to the Free Software Foundation, Inc.,
#' 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.



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


#Write registries to three seperate CSV files
writeToCSV <- function(output_dir=getwd(), current_date=Sys.Date(), index, out_of_date_never_done=data.frame(), at_risk=data.frame(), outliers=data.frame()) {
  file_ending = paste(format(current_date, "_%d%b%Y"), "_", index, ".txt", sep="")
  out_of_date_file = paste("CW_OutOfDate", file_ending, sep="");
  at_risk_file = paste("CW_AtRisk", file_ending, sep="");
  outliers_file = paste("CW_Outliers", file_ending, sep="");
  
  write.csv(out_of_date_never_done,
            file=paste(output_dir, out_of_date_file, sep="/"),
            row.names=FALSE)
  write.csv(at_risk[order(at_risk$Latest.BMI.Percentile, decrease=TRUE), ],
            file=paste(output_dir, at_risk_file, sep="/"),
            row.names=FALSE)
  write.csv(outliers,
            file=paste(output_dir, outliers_file, sep="/"),
            row.names=FALSE)
}

#Write registries to excel file (requires R xlsx package)
writeToExcel <- function(output_dir=getwd(), current_date=Sys.Date(), index, out_of_date_never_done=data.frame(), at_risk=data.frame(), outliers=data.frame()) {
  excel_file = paste(output_dir, 
                     paste("Child_Wellness_Registries_", format(current_date, format="%d%b%Y"), "_", index, ".xlsx", sep=""), 
                     sep="/");
  write.xlsx(out_of_date_never_done,
             file=excel_file, 
             sheetName="Out Of Date", 
             row.names=FALSE);
  write.xlsx(at_risk, 
             file=excel_file, 
             sheetName="At Risk",
             row.names=FALSE,
             append=TRUE);
  write.xlsx(outliers, 
             file=excel_file, 
             sheetName="Outliers",
             row.names=FALSE,
             append=TRUE);
}


runReport <- function() {
  
  #+
  # Prompt user for file to read
  input_files = choose.files();
  
  if (length(input_files) == 0) {
    stop("No file selected");
  }
  
  #+ 
  # Prompt user for folder to save results to 
  output_dir = choose.dir(default=getwd(), caption="Select a directory to save files to");
  
  if (is.na(output_dir)) {
    output_dir = getwd();
    print("Saving to working directory: ");
    print(getwd());
  }
  

  #get age ranges
  minAge = as.numeric(winDialogString(message="What is the minimum age?", default="2"))
  
  if (minAge < 0) {
    minAge = 2
  } else if (minAge > 100) {
    minAge = 100
  } else if (is.na(minAge)) {
    minAge = 2
  }
  
  maxAge = as.numeric(winDialogString(message="What is the maximum age?", default="5"))
  if (maxAge < 0) {
    maxAge = 2
  } else if (maxAge > 100) {
    maxAge = 100
  } else if (is.na(maxAge)) {
    maxAge = 18
  }
   

  master_data = list()
  
  for (i in 1:length(input_files)) {
    
    if (length(input_files) == 1) {
      current_file = input_files
    } else {
      current_file = input_files[i]
    }
        
    #Read Report File
    data = readReport(current_file)
    
    #Handle for empty data (maybe age range doesn't make sense?)
        
    # Get the current date, and one year ago
    current_date = data$Current.Date[1]

    #Get current age in years 
    data$Calc.Age <- (current_date - data$Birth.Date)/365.25
    
    #Create Registries
    reg = getRegistries(data, minAge, maxAge, current_date)
    
    if (length(input_files > 1)) {
      master_data[[i]] = c(reg, current_date)
    }
    
    #Print Graphs
    saveStatusGraph(output_dir, i,
                    nrow(reg$never_done), nrow(reg$up_to_date), nrow(reg$out_of_date), nrow(reg$data),
                    current_date)
    
    saveGrowthConcern(output_dir, i, reg, current_date)
    
    saveHeightWeightCharts(output_dir, i, reg$data)

    #Print Registries
    saveRegistries(output_dir, current_date, i, reg)
  }
  
  #TODO - handle master_data
  if (length(master_data) > 0) {
    createMasterTable(master_data);
  }
  
  winDialog(type="ok",
            paste("Finished! You can find the files in ", output_dir, sep=""));
}
    
readReport <- function(input_file) {
    
    data = read.csv(input_file)
    
    #+
    #' Convert to R dates
    data$Date.of.Latest.Height = as.Date(data$Date.of.Latest.Height, format="%b %d, %Y")
    data$Date.of.Latest.Weight = as.Date(data$Date.of.Latest.Weight, format="%b %d, %Y")
    data$Date.of.Latest.BMI = as.Date(data$Date.of.Latest.BMI, format="%b %d, %Y")
    data$Date.of.Latest.BMI.Percentile = as.Date(data$Date.of.Latest.BMI.Percentile, format="%b %d, %Y")
    data$Current.Date = as.Date(data$Current.Date, format="%b %d, %Y")
    data$Birth.Date = as.Date(data$Birth.Date, format="%b %d, %Y")
        
    # Convert BMI percentile to number
    data$Latest.BMI.Percentile <- as.numeric(as.character(data$Latest.BMI.Percentile))
    data$Latest.BMI <- as.numeric(as.character(data$Latest.BMI))
  
    return(data)
}

getRegistries <- function(data, minAge, maxAge, current_date) {
    #+
    #' Create Registries
    #' May not add up to all patients due to outliers and data entry issues.
    
    #Subset based on user specified age range
    data = subset(data, data$Calc.Age >= minAge & data$Calc.Age <= maxAge);
  
    one_year_ago = seq(current_date, length=2, by= "-12 months")[2]
    
    #' Outliers
    #' BMI < 11 or BMI > 40
    #' Date of measurement more recent than date of report
    #' Height < ?
    #' Weight < ?
  
    #' Remove outliers from dataframe
    outliers = subset(data, data$Latest.BMI < 11 | data$Latest.BMI > 40 | 
                      data$Date.of.Latest.Height > current_date |
                      data$Date.of.Latest.Weight > current_date | 
                      data$Date.of.Latest.BMI > current_date)

    data = data[!data$Patient.. %in% outliers$Patient..,]
        
    out_of_date = subset(data, (data$Date.of.Latest.Height <= one_year_ago | 
                                  data$Date.of.Latest.Weight <= one_year_ago) &
                             !is.na(data$Date.of.Latest.BMI))
    
    never_done = subset(data, is.na(data$Date.of.Latest.BMI))
    
    up_to_date = subset(data, data$Date.of.Latest.Height > one_year_ago &
                            data$Date.of.Latest.Weight > one_year_ago)
    
    #+ Get BMI status of up to date patients
    severely_wasted = subset(up_to_date, 
                             up_to_date$Latest.BMI.Percentile<0.1)
    wasted = subset(up_to_date, 
                    up_to_date$Latest.BMI.Percentile>=0.1 &
                      up_to_date$Latest.BMI.Percentile<3)
    normal = subset(up_to_date, 
                    up_to_date$Latest.BMI.Percentile>3 &
                      up_to_date$Latest.BMI.Percentile<85)
    risk_of_overweight = subset(up_to_date, 
                                up_to_date$Latest.BMI.Percentile>=85 &
                                  up_to_date$Latest.BMI.Percentile<97)
    overweight = subset(up_to_date, 
                        up_to_date$Latest.BMI.Percentile>=97 &
                          up_to_date$Latest.BMI.Percentile<99.9)
    obese = subset(up_to_date, 
                   up_to_date$Latest.BMI.Percentile>=99.9)
    
    #At risk registry is everyone not in the normal weight category
    at_risk = rbind(severely_wasted, 
                    wasted, 
                    risk_of_overweight, 
                    overweight, 
                    obese)
    
    
    registries <- list("never_done" = never_done,
                       "out_of_date" = out_of_date,
                       "up_to_date" = up_to_date,
                       "out_of_date_never_done" = merge(out_of_date, never_done, all=TRUE),
                       "severely_wasted" = severely_wasted,
                       "wasted" = wasted,
                       "normal" = normal,
                       "risk_of_overweight" = risk_of_overweight,
                       "overweight" = overweight,
                       "obese" = obese,
                       "at_risk" = at_risk,
                       "outliers" = outliers,
                       "data" = data)

    return(registries)
}

saveStatusGraph <- function(output_dir, index, 
                            num_never_done, num_up_to_date, num_out_of_date, num_total, 
                            current_date) {
  # Get counts of number of patients in each registry

  status_counts = c(num_total, num_up_to_date, num_out_of_date, num_never_done)
  status_labels = c("Total", "Up to Date", "Out of Date", "Never Done")
  status_colours = c("mediumpurple2", "darkolivegreen3", "orangered3", "dodgerblue3")
  
  filename = paste("BMI_Status", index, "png", sep=".")
  png(filename=paste(output_dir, filename, sep="/"))
  # Make left margin larger for legend text
  par(mar = c(5,8,4,2) + 0.1);
  # Horizontal bar chart
  bp_status <- barplot(status_counts, col=status_colours, horiz=TRUE,
                       legend.text=status_labels,
                       xlab="Number of patients", 
                       main=paste("Total Peds 2 to 5 years (n=",num_total,") \nas of ",
                                  format(current_date, "%b %d, %Y"), sep=""));
  # Add axis labels
  axis(2, at = bp_status, labels=status_labels, las=1);
  text(x=status_counts/2, y=bp_status,
       labels=as.character(status_counts), xpd=TRUE)
  dev.off();
}

saveGrowthConcern <- function(output_dir, index, reg, current_date) {

  # Store results in vectors
  bmi_labels = c("Severely Wasted", "Wasted", "Normal", "Risk of Overweight", "Overweight", "Obese");
  bmi_counts <- c(nrow(reg$severely_wasted), 
                  nrow(reg$wasted), nrow(reg$normal), nrow(reg$risk_of_overweight), nrow(reg$overweight), nrow(reg$obese));
  bmi_colours = c("dodgerblue3", "orangered3", "darkolivegreen3", "mediumpurple2", "mediumturquoise", "orange", "lightskyblue");
  
  filename=paste("BMI_Count", index, "png", sep=".")
  png(filename=paste(output_dir, filename, sep="/"));
  
  # Make left margin larger for legend text
  par(mar = c(5,4,6,2) + 0.1);
  # Plot BMI status
  bp_bmi <- barplot(bmi_counts, 
                    main=paste("Total Peds 2 to 5 years with up to date BMI ",
                               "(n=",nrow(reg$up_to_date),"/", nrow(reg$data), ") \nas of ",
                               format(current_date, "%b %d, %Y"), sep=""), 
                    ylab="Number of patients", 
                    col=bmi_colours,
                    legend.text=bmi_labels,
                    las=2);
  text(bp_bmi, 
       par("usr")[3], 
       labels=bmi_labels, 
       srt=40, xpd=TRUE, adj=c(1.1, 1.1), cex=0.9);
  text(y=bmi_counts/2, x=bp_bmi, 
       labels=as.character(bmi_counts), xpd=TRUE, 
       fontface="bold")
  dev.off();
}

saveHeightWeightCharts <- function(output_dir, index, data) {
  #Plot Date of Latest Height vs Date of Latest Weight
  
  filename = paste("HeightWeightBoxplot", index, "png", sep=".")
  png(filename=paste(output_dir, filename, sep="/"));
  
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

saveRegistries <- function(output_dir, current_date, index, reg) {
      
    #' Prepare to save registries. Check to make sure xlsx library is installed and install if necessary
    #' Write to a CSV text file otherwise.
    if ("xlsx" %in% rownames(installed.packages())) {
      require(xlsx)
      writeToExcel(output_dir, current_date, index, reg$out_of_date_never_done, reg$at_risk, reg$outliers)
    } else {
      
      response = winDialog(type="yesno", 
                           "Did not find Excel libraries. Would you like to install them now?\nYou must be connected to the internet and have Excel installed on this computer")
      
      if (response == "YES") {
        
        install.packages("xlsx")
        require(xlsx)
        
        if ("xlsx" %in% rownames(installed.packages())) {
          writeToExcel(output_dir, current_date, index, reg$out_of_date_never_done, reg$at_risk, reg$outliers)
        } else {
          winDialog(type="ok", 
                    "Something went wrong installing excel libraries ('xlsx'). Writing to text files.");
          writeToCSV(output_dir, current_date, index, reg$out_of_date_never_done, reg$at_risk, reg$outliers)
        }
      }
      else {
        writeToCSV(output_dir, current_date, index, reg$out_of_date_never_done, reg$at_risk, reg$outliers)
      }
  }
}

#TODO - fill this in
createMasterTable <- function(master_data) {
  for (i in length(master_data)) {
    
  }
}
#+
#' Run Application
runReport()