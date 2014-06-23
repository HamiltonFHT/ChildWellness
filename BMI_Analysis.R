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
#' BMI up-to-date
#' Produces two plots:
#'  # of patients in each of the above 3 registries
#'  # of up-to-date patients in each BMI category


writeToCSV <- function(output_dir=getwd(), current_date=Sys.Date(), out_of_date_never_done=data.frame(), at_risk=data.frame(), outliers=data.frame()) {
  file_ending = paste(format(current_date, "_%d%b%Y"), ".txt", sep="")
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

writeToExcel <- function(output_dir=getwd(), current_date=Sys.Date(), out_of_date_never_done=data.frame(), at_risk=data.frame(), outliers=data.frame()) {
  excel_file = paste(output_dir, 
                     paste("Child_Wellness_Registries_", format(current_date, format="%d%b%Y"), ".xlsx", sep=""), 
                     sep="/");
  write.xlsx(out_of_date,
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


#+
# Prompt user for file to read
input_file = file.choose();

if (input_file == "") {
  print("No file selected");
  return;
}
df = read.csv(input_file)

#+ 
# Prompt user for folder to save results to 
output_dir = choose.dir(default=getwd(), caption="Select a directory to save files to");

if (is.na(output_dir)) {
  output_dir = getwd();
  print("Saving to working directory: ");
  print(getwd());
}

#+
#' Convert to R dates
df$Date.of.Latest.Height = as.Date(df$Date.of.Latest.Height, format="%b %d, %Y")
df$Date.of.Latest.Weight = as.Date(df$Date.of.Latest.Weight, format="%b %d, %Y")
df$Date.of.Latest.BMI = as.Date(df$Date.of.Latest.BMI, format="%b %d, %Y")
df$Date.of.Latest.BMI.Percentile = as.Date(df$Date.of.Latest.BMI.Percentile, format="%b %d, %Y")
df$Current.Date = as.Date(df$Current.Date, format="%b %d, %Y")
df$Birth.Date = as.Date(df$Birth.Date, format="%b %d, %Y")

# Get the current date, and one year ago
current_date = df$Current.Date[1]
one_year_ago = seq(current_date, length=2, by= "-12 months")[2]

#Get current age in years 
df$Calc.Age <- (current_date - df$Birth.Date)/365.25

# Filter out patients below the age of 2 and above the age of 5
df = subset(df, df$Calc.Age >= 2 & df$Calc.Age <= 5)

# Convert BMI percentile to number
df$Latest.BMI.Percentile <- as.numeric(as.character(df$Latest.BMI.Percentile))
df$Latest.BMI <- as.numeric(as.character(df$Latest.BMI))

#+
#' Create Registries
#' May not add up to all patients due to outliers and data entry issues.
never_done = subset(df, is.na(df$Date.of.Latest.BMI))
up_to_date = subset(df, df$Date.of.Latest.Height > one_year_ago & df$Date.of.Latest.Weight > one_year_ago)
out_of_date = subset(df, (df$Date.of.Latest.Height <= one_year_ago | 
                            df$Date.of.Latest.Weight <= one_year_ago) &
                       !is.na(df$Date.of.Latest.BMI))
top_85th_percentile = subset(df, df$Latest.BMI.Percentile > 85)

out_of_date_never_done = merge(out_of_date, never_done)


#' Outliers
#' BMI < 11 or BMI > 40
#' Date of measurement more recent than date of report
#' Height < ?
#' Weight < ?

outliers = subset(df, df$Latest.BMI < 11 | df$Latest.BMI > 40 | df$Date.of.Latest.Height > current_date |
                    df$Date.of.Latest.Weight > current_date | df$Date.of.Latest.BMI > current_date)

# Get counts of number of patients in each registry
num_never_done = nrow(never_done)
num_up_to_date = nrow(up_to_date)
num_out_of_date = nrow(out_of_date)
num_total = nrow(df)

status_counts = c(num_total, num_up_to_date, num_out_of_date, num_never_done)
status_labels = c("Total", "Up to Date", "Out of Date", "Never Done")
status_colours = c("mediumpurple2", "darkolivegreen3", "orangered3", "dodgerblue3")

#+
#' Plot number of patients in each of the registries
png(filename=paste(output_dir, "BMI_Status.png", sep="/"))
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
     labels=as.character(status_counts), xpd=TRUE,
     fontface="bold")
dev.off();

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

# Store results in vectors
bmi_labels = c("Severely Wasted", "Wasted", "Normal", "Risk of Overweight", "Overweight", "Obese");
bmi_counts <- c(nrow(severely_wasted), nrow(wasted), nrow(normal), nrow(risk_of_overweight), nrow(overweight), nrow(obese));
bmi_colours = c("dodgerblue3", "orangered3", "darkolivegreen3", "mediumpurple2", "mediumturquoise", "orange", "lightskyblue");

png(filename=paste(output_dir, "BMI_Count.png", sep="/"));
# Make left margin larger for legend text
par(mar = c(5,4,6,2) + 0.1);
# Plot BMI status
bp_bmi <- barplot(bmi_counts, 
                  main=paste("Total Peds 2 to 5 years with up to date BMI ",
                             "(n=",num_up_to_date,"/", num_total, ") \nas of ",
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

HvW_df = df[!df$Patient.. %in% outliers$Patient..,]
png(filename=paste(output_dir, "HeightvsWeight.png", sep="/"));
plot(HvW_df$Date.of.Latest.Weight, HvW_df$Date.of.Latest.Height, 
     xaxt="n", yaxt="n",
     main="Date of Latest Weight vs. Height",
     xlab="Date of Latest Weight", 
     ylab="Date of Latest Height")
axis.Date(side = 2, x=HvW_df$Date.of.Latest.Height, format = "%Y")
axis.Date(side = 1, x=HvW_df$Date.of.Latest.Weight, format = "%Y")
abline(a=0, b=1, col="green")
dev.off()


#' Prepare to save registries. Check to make sure xlsx library is installed and install if necessary
#' Write to a CSV text file otherwise.
if ("xlsx" %in% rownames(installed.packages())) {
  require(xlsx)
  writeToExcel(output_dir, current_date, out_of_date_never_done, at_risk, outliers)
} else {
  
  response = winDialog(type="yesno", 
                       "Did not find Excel libraries. Would you like to install them now?\nYou must be connected to the internet and have Excel installed on this computer")
  
  if (response == "YES") {
    
    install.packages("xlsx")
    
    if ("xlsx" %in% rownames(installed.packages())) {
      writeToExcel(output_dir, current_date, out_of_date_never_done, at_risk, outliers)
    } else {
      winDialog(type="ok", 
                "Something went wrong installing excel libraries ('xlsx'). Writing to text files.");
      writeToCSV(output_dir, current_date, out_of_date_never_done, at_risk, outliers)
    }
  }
  else {
    writeToCSV(output_dir, current_date, out_of_date_never_done, at_risk, outliers)
  }
}

winDialog(type="ok",
          paste("Finished! You can find the files in ", output_dir, sep=""))