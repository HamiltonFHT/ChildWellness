source("../../BMI_Analysis.R")

test_that("Read Data File", {
  input_file = "../../data/ChildWellness_BMI_ExampleDoc_11-Aug-2014.txt"
  
  data <- readReport(input_file)
  
  expect_that(data, is_a("data.frame"))
  expect_that(nrow(data), equals(70))
  
  expect_that(data$Latest.BMI, is_a("numeric"))
  expect_that(data$Date.of.Latest.BMI, is_a("Date"))
  expect_that(data$Current.Date, is_a("Date"))
  
})


test_that("Registries", {
  input_file = "../../data/ChildWellness_BMI_ExampleDoc_11-Aug-2014.txt"
  data <- readReport(input_file)
  
  current_date = data$Current.Date[1]
  expect_that(current_date, is_a("Date"))
  
  data$Calc.Age <- as.numeric((current_date - data$Birth.Date)/365.25)
  expect_that(max(data$Calc.Age), is_less_than(19))
  expect_that(min(data$Calc.Age), is_more_than(2))
  
  reg <- getRegistries(data, 2, 18, current_date)
  expect_that(nrow(reg$data), equals(54))
  expect_that(nrow(reg$outliers), equals(12))
  
  expect_that(reg$severely_wasted, equals(0))
  expect_that(reg$wasted, equals(1))
  expect_that(reg$normal, equals(8))
  expect_that(reg$risk_of_overweight, equals(5))
  expect_that(reg$overweight, equals(3))
  expect_that(reg$obese, equals(1))
   
  expect_that(reg$at_risk, is_a("logical"))
  
  expect_that(sum(reg$at_risk), equals(11))
  expect_that(sum(reg$up_to_date), equals(19))
  expect_that(sum(reg$out_of_date), equals(15))
  expect_that(sum(reg$never_done), equals(20))
  expect_that(sum(reg$out_of_date_never_done), equals(35))
})


