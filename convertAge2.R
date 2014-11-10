convertAge2 <- function(PSS_age) {
  age_in_years = rep(0, length(PSS_age));
  
  for (i in 1:length(PSS_age)) {
    age = PSS_age[i]
    if (length(grep("mo", age)) > 0) { 
      age = as.numeric(gsub("mo", "", age))/12;
    }
    age_in_years[i] = as.numeric(age);
  }
  
  return(age_in_years)
}
