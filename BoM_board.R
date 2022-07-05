library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(officer)
library(janitor)
library(ggtext)
library(simmer)
library(arulesViz)
library(Matrix)
library(palmerpenguins)
library(ggrepel)
library(quantmod)
library(tseries)
library(timeSeries)
library(forecast)
library(xts)
library(flextable)
library(tidyquant)
library(timetk)
library(fs)
library(openxlsx)
library(gt)
library(webshot)
library(plotly)
library(summarytools)
library(questionr)
library(epiDisplay)
library(rmarkdown)
library(DescTools)
library(lubridate)

# Functions
options(scipen = 100000000)


##### Part 1 - Data Reading #####


# AS400<7499> File pulling ---- Change Directory ----

as400_7499 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/BoM Report Automation/Sample for test 2/Book2.xlsx", 
                         col_types = c("numeric", "numeric", "text", 
                                       "text", "numeric", "numeric", "text", 
                                       "text", "numeric", "text", "text", 
                                       "text", "text", "numeric", "numeric", 
                                       "numeric", "numeric", "numeric", 
                                       "numeric", "numeric", "text"))

as400_7499[-1:-5,] -> as400_7499
colnames(as400_7499) <- as400_7499[1, ]
as400_7499[-1, ] -> as400_7499

colnames(as400_7499)[1]  <- "Loc"
colnames(as400_7499)[2]  <- "Product_code"
colnames(as400_7499)[3]  <- "Label"
colnames(as400_7499)[4]  <- "Description"
colnames(as400_7499)[5]  <- "Net_wt"
colnames(as400_7499)[6]  <- "Grs_wt"
colnames(as400_7499)[7]  <- "Formula"
colnames(as400_7499)[8]  <- "Formula_description"
colnames(as400_7499)[9]  <- "Batch_size"
colnames(as400_7499)[10] <- "CompNo_labor_code"
colnames(as400_7499)[11] <- "Comp_description"
colnames(as400_7499)[12] <- "Comp_type"
colnames(as400_7499)[13] <- "Um"
colnames(as400_7499)[14] <- "Required_qty"
colnames(as400_7499)[15] <- "Yield"
colnames(as400_7499)[16] <- "Total_qty"
colnames(as400_7499)[17] <- "Standard_cost"
colnames(as400_7499)[18] <- "Case"
colnames(as400_7499)[19] <- "Pounds"
colnames(as400_7499)[20] <- "Percent_required"
colnames(as400_7499)[21] <- "Xfer_comp_type"




## FG On Hand data pulling ---- Change directory ----

FG_On_Hand <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/BoM Report Automation/Sample for test 2/FG_On_Hand.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "numeric", "numeric", "numeric", 
                                       "numeric", "numeric", "text", "numeric", 
                                       "numeric", "text"))


FG_On_Hand[-1:-3,] -> FG_On_Hand
colnames(FG_On_Hand) <- FG_On_Hand[1, ]
FG_On_Hand[-1, ] -> FG_On_Hand

colnames(FG_On_Hand)[1]  <- "Loc"
colnames(FG_On_Hand)[2]  <- "SKU"
colnames(FG_On_Hand)[3]  <- "Description"
colnames(FG_On_Hand)[4]  <- "On_Hand"
colnames(FG_On_Hand)[5]  <- "Held_Qty"
colnames(FG_On_Hand)[6]  <- "On_Order"
colnames(FG_On_Hand)[7]  <- "Production_Schedule"
colnames(FG_On_Hand)[8]  <- "Standard_Cost"
colnames(FG_On_Hand)[10] <- "First_Prod"
colnames(FG_On_Hand)[11] <- "Open_Order"


FG_On_Hand


## Constrained_RM_List 

Constrained_RM_List <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/Constrained_RM_List.xlsx", 
                                  col_types = c("numeric", "text", "text", 
                                                "numeric", "text"))




## FG_ref_to_mpg_ref 

FG_ref_to_mfg_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/FG_On_Hand/FG_ref_to_mfg_ref.xlsx")




# RM_On_Hand pulling ---- Change Directory ----

RM_On_Hand <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/BoM Report Automation/Sample for test 2/RM_On_Hand.xlsx", 
                         col_types = c("numeric", "numeric", "text", 
                                       "numeric", "numeric", "numeric", 
                                       "numeric", "text", "text", "date"))


RM_On_Hand[-1:-3,] -> RM_On_Hand
colnames(RM_On_Hand) <- RM_On_Hand[1, ]
RM_On_Hand[-1, ] -> RM_On_Hand


colnames(RM_On_Hand)[1]  <- "Loc"
colnames(RM_On_Hand)[2]  <- "Component"
colnames(RM_On_Hand)[3]  <- "Description"
colnames(RM_On_Hand)[4]  <- "On_Hand"
colnames(RM_On_Hand)[5]  <- "Held_Qty"
colnames(RM_On_Hand)[6]  <- "On_Order"
colnames(RM_On_Hand)[7]  <- "Standard_Cost"
colnames(RM_On_Hand)[8]  <- "Um"
colnames(RM_On_Hand)[9]  <- "Class"
colnames(RM_On_Hand)[10] <- "first_PO"

RM_On_Hand %>% 
  dplyr::mutate(First_PO = format(RM_On_Hand$first_PO, format = "%m/%d/%y")) -> RM_On_Hand

RM_On_Hand[,-10] -> RM_On_Hand


# Campus_ref pulling 

Campus_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/RM_on_Hand/Campus_ref.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "numeric"))



colnames(Campus_ref)[1] <- "Loc"
colnames(Campus_ref)[2] <- "Description"
colnames(Campus_ref)[3] <- "Campus_Name"
colnames(Campus_ref)[4] <- "Campus"



# DSX Forecast pulling ---- Change Directory ----

DSX_Forecast_Backup <- read_excel(
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.02.09.xlsx")

DSX_Forecast_Backup[-1,] -> DSX_Forecast_Backup
colnames(DSX_Forecast_Backup) <- DSX_Forecast_Backup[1, ]
DSX_Forecast_Backup[-1, ] -> DSX_Forecast_Backup

colnames(DSX_Forecast_Backup)[1]  <- "Primary_Channel_ID"
colnames(DSX_Forecast_Backup)[2]  <- "Segmentation_ID"
colnames(DSX_Forecast_Backup)[3]  <- "Sub_Segment_ID"
colnames(DSX_Forecast_Backup)[4]  <- "Forecast_Month_Year_Code_Segment_ID"
colnames(DSX_Forecast_Backup)[5]  <- "Product_Manufacturing_Location_Code"
colnames(DSX_Forecast_Backup)[6]  <- "Product_Manufacturing_Location_Name"
colnames(DSX_Forecast_Backup)[7]  <- "Location_No"
colnames(DSX_Forecast_Backup)[8]  <- "Location_Name"
colnames(DSX_Forecast_Backup)[9]  <- "Product_Label_SKU_Code"
colnames(DSX_Forecast_Backup)[10] <- "Product_Label_SKU_Name"
colnames(DSX_Forecast_Backup)[11] <- "Product_Category_Name"
colnames(DSX_Forecast_Backup)[12] <- "Product_Platform_Name"
colnames(DSX_Forecast_Backup)[13] <- "Product_Group_Code"
colnames(DSX_Forecast_Backup)[14] <- "Product_Group_Short_Name"
colnames(DSX_Forecast_Backup)[15] <- "Product_Manufacturing_Line_Area_No_Code"
colnames(DSX_Forecast_Backup)[16] <- "ABC_4_ID"
colnames(DSX_Forecast_Backup)[17] <- "Safety_Stock_ID"
colnames(DSX_Forecast_Backup)[18] <- "MTO_MTS_Gross_Requirements_Calc_Method_ID"
colnames(DSX_Forecast_Backup)[19] <- "Adjusted_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[20] <- "Adjusted_Forecast_Cases"
colnames(DSX_Forecast_Backup)[21] <- "Stat_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[22] <- "Stat_Forecast_Cases"
colnames(DSX_Forecast_Backup)[23] <- "Cust_Ref_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[24] <- "Cust_Ref_Forecast_Cases"

DSX_Forecast_Backup$Forecast_Month_Year_Code_Segment_ID <- as.double(DSX_Forecast_Backup$Forecast_Month_Year_Code_Segment_ID)
DSX_Forecast_Backup$Product_Manufacturing_Location_Code <- as.double(DSX_Forecast_Backup$Product_Manufacturing_Location_Code)
DSX_Forecast_Backup$Location_No <- as.double(DSX_Forecast_Backup$Location_No)
DSX_Forecast_Backup$Product_Manufacturing_Line_Area_No_Code <- as.double(DSX_Forecast_Backup$Product_Manufacturing_Line_Area_No_Code)
DSX_Forecast_Backup$Safety_Stock_ID <- as.double(DSX_Forecast_Backup$Safety_Stock_ID)
DSX_Forecast_Backup$Adjusted_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup$Adjusted_Forecast_Pounds_lbs)
DSX_Forecast_Backup$Adjusted_Forecast_Cases <- as.double(DSX_Forecast_Backup$Adjusted_Forecast_Cases)
DSX_Forecast_Backup$Stat_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup$Stat_Forecast_Pounds_lbs)
DSX_Forecast_Backup$Stat_Forecast_Cases <- as.double(DSX_Forecast_Backup$Stat_Forecast_Cases)
DSX_Forecast_Backup$Cust_Ref_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup$Cust_Ref_Forecast_Pounds_lbs)
DSX_Forecast_Backup$Cust_Ref_Forecast_Cases <- as.double(DSX_Forecast_Backup$Cust_Ref_Forecast_Cases)


DSX_Forecast_Backup


# Open Customer Order File pulling ----  Change Directory ----

Open_Cust_Ord <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/BoM Report Automation/Sample for test 2/wo receipt custord po - 02.09.22.xlsx", 
                            sheet = "custord", col_names = FALSE)

colnames(Open_Cust_Ord)[1] <- "aa"


Open_Cust_Ord %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8", "9", "10")) -> Open_Cust_Ord


Open_Cust_Ord[,c(-1:-2, -5, -7)] -> Open_Cust_Ord

colnames(Open_Cust_Ord)[1]  <- "ProductSkuCode"
colnames(Open_Cust_Ord)[2]  <- "temp_Loc"
colnames(Open_Cust_Ord)[3]  <- "Qty"
colnames(Open_Cust_Ord)[4]  <- "year"
colnames(Open_Cust_Ord)[5]  <- "month"
colnames(Open_Cust_Ord)[6]  <- "day"

Open_Cust_Ord %>% 
  dplyr::mutate(month_year = paste(month,"_",year)) -> Open_Cust_Ord

Open_Cust_Ord %>% 
  dplyr::mutate(month_year = gsub(" ", "", month_year)) -> Open_Cust_Ord

Open_Cust_Ord$temp_Loc -> remove_zero
remove_zero_1 <- sub("^0+", "", remove_zero)
data.frame(remove_zero_1) -> remove_zero_1

cbind(Open_Cust_Ord, remove_zero_1) -> Open_Cust_Ord

Open_Cust_Ord[,-2] -> Open_Cust_Ord
colnames(Open_Cust_Ord)[ncol(Open_Cust_Ord)]  <- "Loc"

Open_Cust_Ord %>% 
  dplyr::relocate(Loc, .before = "Qty") -> Open_Cust_Ord

Open_Cust_Ord %>% 
  dplyr::mutate(ref = paste(Loc, "_", ProductSkuCode)) %>% 
  dplyr::mutate(date = paste(year, "/", month, "/",day)) -> Open_Cust_Ord


Open_Cust_Ord %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) -> Open_Cust_Ord

Open_Cust_Ord %>% 
  dplyr::mutate(date = gsub(" ", "", date)) -> Open_Cust_Ord


Open_Cust_Ord %>% 
  dplyr::relocate(ref, ProductSkuCode) -> Open_Cust_Ord


Open_Cust_Ord %>% 
  dplyr::relocate(date, .after = "Qty") -> Open_Cust_Ord


Open_Cust_Ord$Qty <- as.double(Open_Cust_Ord$Qty)

Open_Cust_Ord %>% 
  dplyr::filter(Loc != "16", Loc != "22", Loc != "502", Loc != "503", Loc != "60S", Loc != "60T") -> Open_Cust_Ord


### Need to delete two locations (16, 22, 502, 503, 060S, 060T)


###############################################################################################################################################################
###############################################################################################################################################################
###############################################################################################################################################################
###############################################################################################################################################################
###############################################################################################################################################################
###############################################################################################################################################################


##### Part 2 - Data Wrangling #####


# Data wrangling - as400_7499 part 1 

as400_7499 %>% 
  dplyr::filter(!is.na(Label)) -> as400_7499

as400_7499 %>%  
  dplyr::filter(Um == "EA" | Um == "LB") -> as400_7499

# Data wrangling - as400_7499 part 2 

as400_7499 %>% 
  dplyr::mutate(ref = paste(Loc, "_", Product_code, Label)) %>% 
  dplyr::mutate(component_ref = paste(Loc, "_", CompNo_labor_code)) %>% 
  dplyr::mutate(SKU = paste(Product_code, Label)) %>%
  dplyr::mutate(Constrained_only = "") %>% 
  dplyr::relocate(ref, component_ref, SKU, Constrained_only) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(component_ref = gsub(" ", "", component_ref)) -> as400_7499

# Constrained_RM_List for vlookup 

colnames(Constrained_RM_List)[1]  <- "CompNo_labor_code"
colnames(Constrained_RM_List)[2]  <- "Type"
colnames(Constrained_RM_List)[3]  <- "Starch_Allocation"

Constrained_RM_List


merge(as400_7499, Constrained_RM_List[, c("CompNo_labor_code", "Type")], by = "CompNo_labor_code", all.x = TRUE) -> as400_7499


as400_7499[,-5] -> as400_7499

as400_7499 %>% 
  dplyr::relocate(ref, component_ref, SKU, Type, Loc, Product_code) -> as400_7499

colnames(as400_7499)[4]  <- "Constrained_only"




# Data wrangling - FG_On_Hand 

FG_On_Hand %>% 
  dplyr::mutate(ref = paste(Loc, "_", SKU)) %>% 
  dplyr::relocate(ref) -> FG_On_Hand


FG_On_Hand %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) -> FG_On_Hand

merge(FG_On_Hand, FG_ref_to_mfg_ref[, c("ref", "Mfg_Ref")], by = "ref", all.x = TRUE) -> FG_On_Hand

FG_On_Hand %>% 
  dplyr::relocate(ref, Mfg_Ref) -> FG_On_Hand


FG_On_Hand %>% 
  dplyr::mutate(Mfg_Ref = gsub(" ", "", Mfg_Ref)) -> FG_On_Hand

colnames(FG_On_Hand)[2]  <- "mpg_ref"



# get Pivot table for FG_On_Hand 

reshape2::dcast(FG_On_Hand, mpg_ref ~ . , value.var = "On_Hand", sum) -> FG_On_Hand_pivot

colnames(FG_On_Hand_pivot)[1]  <- "ref"
colnames(FG_On_Hand_pivot)[2]  <- "sum_of_On_Hand"

FG_On_Hand_pivot %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) -> FG_On_Hand_pivot

# Data wrangling FG_On_Hand_Pivot

merge(as400_7499, FG_On_Hand_pivot[, c("ref", "sum_of_On_Hand")], by = "ref", all.x = TRUE) -> as400_7499

as400_7499 %>% 
  dplyr::relocate(sum_of_On_Hand, .after = Description) -> as400_7499

colnames(as400_7499)[10]  <- "FG_On_hand"



# Data wrangling - RM_On_Hand 

merge(RM_On_Hand, Campus_ref[, c("Loc", "Campus")], by = "Loc", all.x = TRUE) %>% 
  dplyr::mutate(ref = paste(Loc, "_", Component)) -> RM_On_Hand

RM_On_Hand %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) -> RM_On_Hand

RM_On_Hand %>% 
  dplyr::mutate(campus_ref = paste(Campus, "_" ,Component)) %>% 
  dplyr::relocate(ref, campus_ref, Campus) -> RM_On_Hand

RM_On_Hand %>% 
  dplyr::mutate(campus_ref = gsub(" ", "", campus_ref)) -> RM_On_Hand

reshape2::dcast(RM_On_Hand, campus_ref ~ . , value.var = "On_Hand", sum) -> RM_On_Hand_Pivot


colnames(RM_On_Hand_Pivot)[1]  <- "component_ref"
colnames(RM_On_Hand_Pivot)[2]  <- "RM_On_Hand"


merge(as400_7499, RM_On_Hand_Pivot[, c("component_ref", "RM_On_Hand")], by = "component_ref", all.x = TRUE) -> as400_7499

as400_7499 %>% 
  dplyr::relocate(RM_On_Hand, .after = "Comp_description") -> as400_7499

as400_7499$RM_On_Hand -> rmOnHand_1

rmOnHand_1[is.na(rmOnHand_1)] <- 0

cbind(as400_7499, rmOnHand_1) -> as400_7499

colnames(as400_7499)[ncol(as400_7499)] <- "RM_on_Hand"

as400_7499 %>% 
  dplyr::relocate(RM_on_Hand, .after = "Comp_description") -> as400_7499

as400_7499[, -18] -> as400_7499
colnames(as400_7499)[17] <- "RM_On_Hand"

# Data wrangling - DSX_Forecast_Backup 

DSX_Forecast_Backup %>% 
  dplyr::mutate(Product_Label_SKU = gsub("-", "", Product_Label_SKU_Code)) -> DSX_Forecast_Backup

DSX_Forecast_Backup[,-9] -> DSX_Forecast_Backup

DSX_Forecast_Backup %>% 
  dplyr::relocate(Product_Label_SKU, .after = "Location_Name") -> DSX_Forecast_Backup


DSX_Forecast_Backup %>% 
  dplyr::mutate(ref = paste(Location_No, "_", Product_Label_SKU)) %>% 
  dplyr::mutate(mfg_ref = paste(Product_Manufacturing_Location_Code, "_", Product_Label_SKU)) %>% 
  dplyr::relocate(ref, mfg_ref) -> DSX_Forecast_Backup

DSX_Forecast_Backup %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) -> DSX_Forecast_Backup

DSX_Forecast_Backup %>% 
  dplyr::mutate(mfg_ref = gsub(" ", "", mfg_ref)) -> DSX_Forecast_Backup

DSX_Forecast_Backup$Forecast_Month_Year_Code_Segment_ID <- as.character(DSX_Forecast_Backup$Forecast_Month_Year_Code_Segment_ID)
DSX_Forecast_Backup[is.na(DSX_Forecast_Backup)] <- 0


# value n/a to 0
DSX_Forecast_Backup$Adjusted_Forecast_Pounds_lbs -> dsx_na_1
DSX_Forecast_Backup$Adjusted_Forecast_Cases      -> dsx_na_2
DSX_Forecast_Backup$Stat_Forecast_Pounds_lbs     -> dsx_na_3
DSX_Forecast_Backup$Stat_Forecast_Cases          -> dsx_na_4
DSX_Forecast_Backup$Cust_Ref_Forecast_Pounds_lbs -> dsx_na_5
DSX_Forecast_Backup$Cust_Ref_Forecast_Cases      -> dsx_na_6


dsx_na_1[is.na(dsx_na_1)] <- 0
dsx_na_2[is.na(dsx_na_2)] <- 0
dsx_na_3[is.na(dsx_na_3)] <- 0
dsx_na_4[is.na(dsx_na_4)] <- 0
dsx_na_5[is.na(dsx_na_5)] <- 0
dsx_na_6[is.na(dsx_na_6)] <- 0

cbind(DSX_Forecast_Backup, dsx_na_1, dsx_na_2, dsx_na_3, dsx_na_4, dsx_na_5, dsx_na_6) -> DSX_Forecast_Backup

DSX_Forecast_Backup[, -21:-26] -> DSX_Forecast_Backup

colnames(DSX_Forecast_Backup)[21] <- "Adjusted_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[22] <- "Adjusted_Forecast_Cases"
colnames(DSX_Forecast_Backup)[23] <- "Stat_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[24] <- "Stat_Forecast_Cases"
colnames(DSX_Forecast_Backup)[25] <- "Cust_Ref_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[26] <- "Cust_Ref_Forecast_Cases"


reshape2::dcast(DSX_Forecast_Backup, mfg_ref ~ Forecast_Month_Year_Code_Segment_ID , value.var = "Adjusted_Forecast_Cases", sum) -> DSX_pivot_1



colnames(DSX_pivot_1)[1]  <- "ref"

colnames(DSX_pivot_1)[3]  <- "Mon_a_fcst"
colnames(DSX_pivot_1)[4]  <- "Mon_b_fcst"
colnames(DSX_pivot_1)[5]  <- "Mon_c_fcst"
colnames(DSX_pivot_1)[6]  <- "Mon_d_fcst"
colnames(DSX_pivot_1)[7]  <- "Mon_e_fcst"
colnames(DSX_pivot_1)[8]  <- "Mon_f_fcst"
colnames(DSX_pivot_1)[9]  <- "Mon_g_fcst"



# as400_7499 & DSX_pivot_1 merge

merge(as400_7499, DSX_pivot_1[, c("ref", "Mon_a_fcst")], by = "ref", all.x = TRUE) -> as400_7499
merge(as400_7499, DSX_pivot_1[, c("ref", "Mon_b_fcst")], by = "ref", all.x = TRUE) -> as400_7499
merge(as400_7499, DSX_pivot_1[, c("ref", "Mon_c_fcst")], by = "ref", all.x = TRUE) -> as400_7499
merge(as400_7499, DSX_pivot_1[, c("ref", "Mon_d_fcst")], by = "ref", all.x = TRUE) -> as400_7499
merge(as400_7499, DSX_pivot_1[, c("ref", "Mon_e_fcst")], by = "ref", all.x = TRUE) -> as400_7499
merge(as400_7499, DSX_pivot_1[, c("ref", "Mon_f_fcst")], by = "ref", all.x = TRUE) -> as400_7499
merge(as400_7499, DSX_pivot_1[, c("ref", "Mon_g_fcst")], by = "ref", all.x = TRUE) -> as400_7499



# as400_7499 "na" to 0

as400_7499$Mon_a_fcst[is.na(as400_7499$Mon_a_fcst)] <- 0
as400_7499$Mon_b_fcst[is.na(as400_7499$Mon_b_fcst)] <- 0
as400_7499$Mon_c_fcst[is.na(as400_7499$Mon_c_fcst)] <- 0
as400_7499$Mon_d_fcst[is.na(as400_7499$Mon_d_fcst)] <- 0
as400_7499$Mon_e_fcst[is.na(as400_7499$Mon_e_fcst)] <- 0
as400_7499$Mon_f_fcst[is.na(as400_7499$Mon_f_fcst)] <- 0
as400_7499$Mon_g_fcst[is.na(as400_7499$Mon_g_fcst)] <- 0



# Data wrangling - Open_Cust_Ord 
merge(Open_Cust_Ord, FG_ref_to_mfg_ref[, c("ref", "Mfg_Ref")], by = "ref", all.x = TRUE) -> Open_Cust_Ord

Open_Cust_Ord %>% 
  dplyr::mutate(Mfg_Ref = gsub(" ", "", Mfg_Ref)) -> Open_Cust_Ord

colnames(Open_Cust_Ord)[1] <- "ref.1"
colnames(Open_Cust_Ord)[10] <- "ref"


Open_Cust_Ord$date <- as.Date(Open_Cust_Ord$date)

Open_Cust_Ord %>% 
  dplyr::mutate(next_30_days = ifelse(date < Sys.Date() + 30, "Y", "N")) -> Open_Cust_Ord

reshape2::dcast(Open_Cust_Ord, ref ~ next_30_days, value.var = "Qty", sum) -> Open_Cust_Ord_Pivot

merge(as400_7499, Open_Cust_Ord_Pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) -> as400_7499



as400_7499$Y -> Y_temp
Y_temp[is.na(Y_temp)] <- 0

cbind(as400_7499, Y_temp) -> as400_7499
as400_7499[, -(ncol(as400_7499)-1)] -> as400_7499

colnames(as400_7499)[ncol(as400_7499)] <- "Open_CustOrd"

as400_7499 %>% 
  dplyr::relocate(Open_CustOrd, .after = Xfer_comp_type) -> as400_7499

as400_7499$Open_CustOrd[is.na(as400_7499$Open_CustOrd)] <- 0


# RM_Mon_a_dep_demand and other dep_demand

as400_7499[, "RM_Mon_a_dep_demand_max"] <- apply(as400_7499[, c("Open_CustOrd", "Mon_a_fcst")], 1, max)

as400_7499 %>% 
  dplyr::mutate(RM_Mon_a_dep_demand = RM_Mon_a_dep_demand_max * Total_qty) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(RM_Mon_b_dep_demand = Mon_b_fcst * Total_qty) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(RM_Mon_c_dep_demand = Mon_c_fcst * Total_qty) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(RM_Mon_d_dep_demand = Mon_d_fcst * Total_qty) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(RM_Mon_e_dep_demand = Mon_e_fcst * Total_qty) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(RM_Mon_f_dep_demand = Mon_f_fcst * Total_qty) -> as400_7499

as400_7499 %>% 
  dplyr::mutate(RM_Mon_g_dep_demand = Mon_g_fcst * Total_qty) -> as400_7499

as400_7499 %>% 
  dplyr::relocate(RM_Mon_a_dep_demand_max) -> as400_7499

as400_7499[, -1] -> as400_7499


# Create weeks_on_hand dataset

reshape2::dcast(as400_7499, Loc + ref + Description + FG_On_hand + Mon_a_fcst + Mon_b_fcst + Mon_c_fcst ~ .) -> weeks_on_hand


weeks_on_hand[, -ncol(weeks_on_hand)] -> weeks_on_hand

weeks_on_hand[, "sum_3monts"] <- apply(weeks_on_hand[, c("Mon_a_fcst", "Mon_b_fcst", "Mon_c_fcst")], 1, sum)





data.frame(round(weeks_on_hand$sum_3monts)) -> sum_3

cbind(weeks_on_hand, sum_3) -> weeks_on_hand
weeks_on_hand[, -8] -> weeks_on_hand
colnames(weeks_on_hand)[8] <- "sum_3"


weeks_on_hand %>% 
  dplyr::mutate(FG_weeks_on_hand = FG_On_hand/(sum_3/13)) -> weeks_on_hand


# Inf/NaN/NA resolve

weeks_on_hand$FG_weeks_on_hand -> FG_inf

replace(FG_inf, is.infinite(FG_inf) | is.nan(FG_inf) | is.na(FG_inf), 0) -> FG_inf


cbind(weeks_on_hand, FG_inf) -> weeks_on_hand


weeks_on_hand$FG_weeks_on_hand[is.na(weeks_on_hand$FG_weeks_on_hand)] <- 0

weeks_on_hand[, -9] -> weeks_on_hand

colnames(weeks_on_hand)[9] <- "FG_weeks_on_hand"



# merge with as400_7499 for FG_weeks_on_hand

merge(as400_7499, weeks_on_hand[, c("ref", "FG_weeks_on_hand")], by = "ref", all.x = TRUE) -> as400_7499


as400_7499 %>% 
  dplyr::relocate(FG_weeks_on_hand, .after = FG_On_hand) -> as400_7499

# Mon_b_fcst_in_lbs.

as400_7499 %>% 
  dplyr::mutate(Mon_b_fcst_in_lbs = Net_wt * Mon_b_fcst) -> as400_7499

as400_7499 %>% 
  dplyr::relocate(Mon_b_fcst_in_lbs, .before = Net_wt) -> as400_7499


# New pivot for another weeks_on_hand

reshape2::dcast(as400_7499, Loc + component_ref + Comp_description + RM_On_Hand  ~ ., value.var = "RM_Mon_a_dep_demand", sum) -> a
reshape2::dcast(as400_7499, Loc + component_ref + Comp_description + RM_On_Hand  ~ ., value.var = "RM_Mon_b_dep_demand", sum) -> b
reshape2::dcast(as400_7499, Loc + component_ref + Comp_description + RM_On_Hand  ~ ., value.var = "RM_Mon_c_dep_demand", sum) -> c



a[, 1:4] -> abc_base

a$. -> a
sprintf('%.2f', a) -> a
data.frame(a) -> a


b$. -> b
sprintf('%.2f', b) -> b
data.frame(b) -> b



c$. -> c
sprintf('%.2f', c) -> c
data.frame(c) -> c


cbind(abc_base, a, b, c) -> weeks_on_hand_2

colnames(weeks_on_hand_2)[5] <- "sum_RM_Mon_a_dep_demand"
colnames(weeks_on_hand_2)[6] <- "sum_RM_Mon_b_dep_demand"
colnames(weeks_on_hand_2)[7] <- "sum_RM_Mon_c_dep_demand"

weeks_on_hand_2$sum_RM_Mon_a_dep_demand <- as.double(weeks_on_hand_2$sum_RM_Mon_a_dep_demand)
weeks_on_hand_2$sum_RM_Mon_b_dep_demand <- as.double(weeks_on_hand_2$sum_RM_Mon_b_dep_demand)
weeks_on_hand_2$sum_RM_Mon_c_dep_demand <- as.double(weeks_on_hand_2$sum_RM_Mon_c_dep_demand)

weeks_on_hand_2 %>% 
  dplyr::mutate(sum_3months = sum_RM_Mon_a_dep_demand + sum_RM_Mon_b_dep_demand + sum_RM_Mon_c_dep_demand) -> weeks_on_hand_2

# weeks_on_hand_2[, "sum"] <- apply(weeks_on_hand_2[, c("sum_RM_Mon_a_dep_demand", "sum_RM_Mon_b_dep_demand", "sum_RM_Mon_c_dep_demand")], 1, sum)

weeks_on_hand_2 %>% 
  dplyr::mutate(RM_weeks_on_hand = RM_On_Hand/(sum_3months/13)) -> weeks_on_hand_2

weeks_on_hand_2$RM_weeks_on_hand  -> rm_decimal_1

sprintf('%.1f', rm_decimal_1) -> rm_decimal_1
cbind(weeks_on_hand_2, rm_decimal_1) -> weeks_on_hand_2

weeks_on_hand_2[, -9] -> weeks_on_hand_2

weeks_on_hand_2$rm_decimal_1 <- as.double(weeks_on_hand_2$rm_decimal_1)
colnames(weeks_on_hand_2)[9] <- "RM_weeks_on_hand"

# Inf resolve

weeks_on_hand_2$RM_weeks_on_hand -> FG_inf_2

replace(FG_inf_2, is.infinite(FG_inf_2) | is.nan(FG_inf_2) | is.na(FG_inf_2), 0) -> FG_inf_2


cbind(weeks_on_hand_2, FG_inf_2) -> weeks_on_hand_2


weeks_on_hand_2$RM_weeks_on_hand[is.na(weeks_on_hand_2$RM_weeks_on_hand)] <- 0

weeks_on_hand_2[, -9] -> weeks_on_hand_2

colnames(weeks_on_hand_2)[9] <- "RM_weeks_on_hand"


merge(as400_7499, weeks_on_hand_2[, c("component_ref", "RM_weeks_on_hand")], by = "component_ref", all.x = TRUE) -> as400_7499

as400_7499$RM_weeks_on_hand[is.na(as400_7499$RM_weeks_on_hand)] <- 0

as400_7499 %>% 
  dplyr::relocate(RM_weeks_on_hand, .after = RM_On_Hand) -> as400_7499



# Final touch
as400_7499 %>% 
  dplyr::relocate(ref) -> as400_7499 

# export it ----

writexl::write_xlsx(as400_7499,
                    path = "C:/Users/SLee/OneDrive - Ventura Foods/Desktop/BoM_Report_2.9.2022.xlsx")


