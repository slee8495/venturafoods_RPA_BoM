library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(officer)
library(openxlsx)
library(lubridate)
library(magrittr)
library(skimr)


##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################


## FG_ref_to_mpg_ref 

FG_ref_to_mfg_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/FG_on_Hand/FG_ref_to_mfg_ref.xlsx")

FG_ref_to_mfg_ref %<>% 
  dplyr::mutate(ref = gsub("-", "_", ref),
                Campus_Ref = gsub("-", "_", Campus_Ref),
                Mfg_Ref = gsub("-", "_", Mfg_Ref)) %>% 
  dplyr::rename(campus_ref = Campus_Ref,
                mfg_ref = Mfg_Ref)




# Campus_ref pulling 

Campus_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/RM_on_Hand/Campus_ref.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "numeric"))



colnames(Campus_ref)[1] <- "Loc"
colnames(Campus_ref)[2] <- "Description"
colnames(Campus_ref)[3] <- "Campus_Name"
colnames(Campus_ref)[4] <- "Campus"

# (Path revision needed) Category (From BI) ---- 
category_bi <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/BI Category and Platform and pack size.xlsx")

category_bi[-1, ] -> category_bi
colnames(category_bi) <- category_bi[1, ]
category_bi[-1, ] -> category_bi

category_bi %>% 
  dplyr::select(1, 3, 6) %>% 
  dplyr::rename(Item = "SKU Code",
                Category = "Product Category Name",
                Platform = "Product Platform Description") %>% 
  dplyr::mutate(Item = gsub("-", "", Item)) -> category_bi


# (Path revision needed) Inventory Model  (Make sure to remove the password of the original .xlsx file) ----
# Make sure with the password

inventory_model <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/Desktop/SS Optimization by Location - Finished Goods October 2022.xlsx",
                              col_names = FALSE, sheet = "Fin Goods")

inventory_model[-1:-7, ] -> inventory_model
colnames(inventory_model) <- inventory_model[1, ]
inventory_model[-1, ] -> inventory_model

inventory_model %>% 
  dplyr::select("Ship Ref", "Net Wt") %>% 
  dplyr::rename(ref = "Ship Ref",
                Net_wt = "Net Wt") %>% 
  dplyr::mutate(ref = gsub("-", "_", ref),
                Net_wt = as.numeric(Net_wt)) -> inventory_model

# (Path revision needed) IOM MicroStrategy ----
IOM_micro <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/IOM Data Extract - 10.19.22.xlsx")

IOM_micro[-1, ] -> IOM_micro
colnames(IOM_micro) <- IOM_micro[1, ]
IOM_micro[-1, ] -> IOM_micro

IOM_micro %>% 
  dplyr::select("Product Label (SKU)", "FG Net Weight") %>% 
  dplyr::rename(Parent_Item_Number = "Product Label (SKU)",
                Net_wt = "FG Net Weight") %>% 
  dplyr::mutate(Net_wt = as.numeric(Net_wt),
                Parent_Item_Number = gsub("-", "", Parent_Item_Number)) -> IOM_micro




# (Path revision needed) DSX Forecast backup ----

DSX_Forecast_Backup <- read_excel(
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/DSX Forecast Backup - 2022.10.19.xlsx")

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


type_convert(DSX_Forecast_Backup) -> DSX_Forecast_Backup

DSX_Forecast_Backup %>% 
  data.frame() -> DSX_Forecast_Backup

DSX_Forecast_Backup %>% 
  dplyr::mutate(Product_Label_SKU_Code = gsub("-", "", Product_Label_SKU_Code)) %>% 
  dplyr::mutate(ref = paste0(Location_No, "_", Product_Label_SKU_Code)) %>% 
  dplyr::mutate(mfg_ref = paste0(Product_Manufacturing_Location_Code, "_", Product_Label_SKU_Code)) %>% 
  dplyr::relocate(ref, mfg_ref) %>% 
  dplyr::mutate(Adjusted_Forecast_Pounds_lbs = replace(Adjusted_Forecast_Pounds_lbs, is.na(Adjusted_Forecast_Pounds_lbs), 0),
                Adjusted_Forecast_Cases = replace(Adjusted_Forecast_Cases, is.na(Adjusted_Forecast_Cases), 0),
                Stat_Forecast_Pounds_lbs = replace(Stat_Forecast_Pounds_lbs, is.na(Stat_Forecast_Pounds_lbs), 0),
                Stat_Forecast_Cases = replace(Stat_Forecast_Cases, is.na(Stat_Forecast_Cases), 0),
                Cust_Ref_Forecast_Pounds_lbs = replace(Cust_Ref_Forecast_Pounds_lbs, is.na(Cust_Ref_Forecast_Pounds_lbs), 0),
                Cust_Ref_Forecast_Cases = replace(Cust_Ref_Forecast_Cases, is.na(Cust_Ref_Forecast_Cases), 0)) -> DSX_Forecast_Backup


# DSX Pivot table
reshape2::dcast(DSX_Forecast_Backup, mfg_ref ~ Forecast_Month_Year_Code_Segment_ID , value.var = "Adjusted_Forecast_Cases", sum) -> DSX_pivot_1



colnames(DSX_pivot_1)[1]  <- "ref"
colnames(DSX_pivot_1)[2]  <- "last_mon_fcst"
colnames(DSX_pivot_1)[3]  <- "Mon_a_fcst"
colnames(DSX_pivot_1)[4]  <- "Mon_b_fcst"
colnames(DSX_pivot_1)[5]  <- "Mon_c_fcst"
colnames(DSX_pivot_1)[6]  <- "Mon_d_fcst"
colnames(DSX_pivot_1)[7]  <- "Mon_e_fcst"
colnames(DSX_pivot_1)[8]  <- "Mon_f_fcst"
colnames(DSX_pivot_1)[9]  <- "Mon_g_fcst"
colnames(DSX_pivot_1)[10]  <- "Mon_h_fcst"
colnames(DSX_pivot_1)[11]  <- "Mon_i_fcst"
colnames(DSX_pivot_1)[12]  <- "Mon_j_fcst"
colnames(DSX_pivot_1)[13]  <- "Mon_k_fcst"
colnames(DSX_pivot_1)[14]  <- "Mon_l_fcst"
colnames(DSX_pivot_1)[15]  <- "Mon_m_fcst"


DSX_pivot_1 %>% 
  dplyr::mutate(last_mon_fcst = round(last_mon_fcst, 0),
                Mon_a_fcst = round(Mon_a_fcst, 0),
                Mon_b_fcst = round(Mon_b_fcst, 0),
                Mon_c_fcst = round(Mon_c_fcst, 0),
                Mon_d_fcst = round(Mon_d_fcst, 0),
                Mon_e_fcst = round(Mon_e_fcst, 0),
                Mon_f_fcst = round(Mon_f_fcst, 0),
                Mon_g_fcst = round(Mon_g_fcst, 0),
                Mon_h_fcst = round(Mon_h_fcst, 0),
                Mon_i_fcst = round(Mon_i_fcst, 0),
                Mon_j_fcst = round(Mon_j_fcst, 0),
                Mon_k_fcst = round(Mon_k_fcst, 0),
                Mon_l_fcst = round(Mon_l_fcst, 0),
                Mon_m_fcst = round(Mon_m_fcst, 0))  -> DSX_pivot_1



# (Path revision needed) Opencustord ----

Open_Cust_Ord <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/wo receipt custord po - 10.19.22.xlsx", 
                            sheet = "custord", col_names = FALSE)

Open_Cust_Ord %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8", "9"), sep = "~") %>% 
  dplyr::select(-"3", -"4", -"6", -"7", -"8") %>% 
  dplyr::rename(aa = "1") %>% 
  tidyr::separate(aa, c("global", "rp", "ProductSkuCode")) %>% 
  dplyr::select(-global, -rp) %>% 
  dplyr::rename(Loc = "2", 
                Qty = "5",
                date = "9") %>% 
  dplyr::mutate(Loc = sub("^0+", "", Loc)) %>% 
  dplyr::mutate(ref = paste0(Loc, "_", ProductSkuCode)) %>% 
  dplyr::mutate(year = lubridate::year(date), year = as.character(year),
                month = lubridate::month(date), month = as.character(month),
                day = lubridate::day(date), day = as.character(day),
                month_year = paste0(month, "_", year),
                Qty = as.double(Qty)) %>% 
  dplyr::filter(Loc != "16", Loc != "22", Loc != "502", Loc != "503", Loc != "60S", Loc != "60T") %>% 
  dplyr::relocate(ref, ProductSkuCode, Loc, Qty, date, year, month, day, month_year) %>% 
  dplyr::mutate(date = as.Date(date)) -> Open_Cust_Ord



# (Path revision needed) Sales and Open orders cube from Micro (Canada only) ----

canada_micro <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/Canada open order - 10.19.22.xlsx", 
                           col_names = FALSE)


canada_micro[-1:-2, ] -> canada_micro
colnames(canada_micro) <- canada_micro[1, ]
canada_micro[-1, ] -> canada_micro

colnames(canada_micro)[1] <- "Loc"
colnames(canada_micro)[2] <- "Loc_Name"
colnames(canada_micro)[3] <- "ProductSkuCode"
colnames(canada_micro)[4] <- "Description"
colnames(canada_micro)[5] <- "date"
colnames(canada_micro)[6] <- "Qty"

canada_micro %>% 
  dplyr::mutate(date = as.integer(date),
                date = as.Date(date, origin = "1899-12-30")) %>% 
  dplyr::select(-Loc_Name, - Description) %>% 
  dplyr::mutate(year = lubridate::year(date), year = as.character(year),
                month = lubridate::month(date), month = as.character(month),
                day = lubridate::day(date), day = as.character(day),
                month_year = paste0(month, "_", year),
                Qty = as.double(Qty),
                ProductSkuCode = gsub("-", "", ProductSkuCode),
                ref = paste0(Loc, "_", ProductSkuCode)) %>% 
  dplyr::relocate(ref, ProductSkuCode, Loc, Qty) -> canada_micro


# combine Open orders and Canada from micro

rbind(Open_Cust_Ord, canada_micro) -> Open_Cust_Ord

merge(Open_Cust_Ord, FG_ref_to_mfg_ref[, c("ref", "mfg_ref")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(next_28_days = ifelse(date <= Sys.Date() + 28, "Y", "N")) -> Open_Cust_Ord

reshape2::dcast(Open_Cust_Ord, ref ~ next_28_days, value.var = "Qty", sum) -> Open_Cust_Ord_Pivot



# (Path revision needed) Read JDE BoM ----

jde_bom <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/JDE BoM 10.19.22.xlsx", 
                      col_names = FALSE)


jde_bom[-1:-2, ] -> jde_bom
colnames(jde_bom) <- jde_bom[1, ] 
jde_bom[-1, ] -> jde_bom
jde_bom[ , c(-4,-16:-19)] -> jde_bom

names(jde_bom) <- stringr::str_replace_all(names(jde_bom), c(" " = "_"))
type_convert(jde_bom) -> jde_bom


jde_bom %<>% 
  dplyr::mutate(ref = paste0(Business_Unit, "_", Parent_Item_Number),
                comp_ref = paste0(Business_Unit, "_", Component))

colnames(jde_bom)[13] <- "Quantity_w_Scrap"


# (Path revision needed) AS400-86 ----

as400_86 <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/AS400 loc 86 BoM.xlsx", 
                       col_names = FALSE)



as400_86[-1:-6, ] -> as400_86
colnames(as400_86) <- as400_86[1, ]
as400_86[-1, ] -> as400_86

type_convert(as400_86) -> as400_86
names(as400_86) <- stringr::str_replace_all(names(as400_86), c(" " = "_"))

as400_86 %>% filter(!is.na(Label))

as400_86 %>% 
  dplyr::rename(Business_Unit = "Loc",
                Parent_Description = "Description",
                Component = "Comp#/labor_code",
                Component_Description = "Comp_description",
                Commodity_Class = "Comp_type",
                UM = "Um",
                Percent_Scrap = "Yield",
                Quantity_Per = "Required_qty",
                "Quantity_w/_Scrap" = "Total_qty",
                Unit_Cost = "Standard_cost") %>% 
  dplyr::filter(!is.na(Label)) %>%
  dplyr::filter(!is.na(UM)) %>% 
  dplyr::filter(UM == "LB" | UM == "EA") %>% 
  dplyr::mutate(Parent_Item_Number = paste0(Product_code, Label),
                Business_Unit = as.integer(Business_Unit),
                ref = paste0(Business_Unit, "_", Parent_Item_Number),
                comp_ref = paste0(Business_Unit, "_", Component)) %>% 
  dplyr::select(-Product_code, -Label, -Net_wt, -Grs_wt, -Formula, -Formula_description, -Batch_size,
                -Case, -Pounds, -Percent_required, -Xfer_comp_type) %>% 
  dplyr::mutate(Level = "", UOM = "", Stocking_Type = "", ) %>% 
  dplyr::relocate(Business_Unit, Level, Parent_Item_Number, Parent_Description,	UOM, Component,	
                  Component_Description, Commodity_Class,	UM,	Stocking_Type, Percent_Scrap,	Quantity_Per,	
                  "Quantity_w/_Scrap",	Unit_Cost, ref, comp_ref) -> as400_86

colnames(as400_86)[13] <- "Quantity_w_Scrap"


rbind(jde_bom, as400_86) -> jde_bom

jde_bom %<>% 
  dplyr::mutate(Component = replace(Component, is.na(Component), "NA"))

# parent count
jde_bom %>% 
  dplyr::group_by(comp_ref, Parent_Item_Number) %>% 
  dplyr::summarize(count = n()) %>% 
  dplyr::mutate(parent_count_1 = table(comp_ref)) %>% 
  dplyr::mutate(parent_count_1 = as.integer(parent_count_1)) %>% 
  dplyr::select(-count, -Parent_Item_Number) -> parent_count_1

parent_count_1[-which(duplicated(parent_count_1$comp_ref)),] -> parent_count_1


jde_bom %>% 
  dplyr::group_by(Component, Parent_Item_Number) %>% 
  dplyr::summarize(count = n()) %>% 
  dplyr::mutate(parent_count_2 = table(Component)) %>% 
  dplyr::mutate(parent_count_2 = as.integer(parent_count_2)) %>% 
  dplyr::select(-count, -Parent_Item_Number) -> parent_count_2

parent_count_2[-which(duplicated(parent_count_2$Component)),] -> parent_count_2



# (Path revision needed) Inventory from MicroStrategy (FG) ----

FG <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/Inventory Report for all locations - 10.19.22.xlsx", 
                 col_names = FALSE,
                 sheet = "FG")


FG[-1:-2, ] -> FG
colnames(FG) <- FG[1, ]
FG[-1, ] -> FG

colnames(FG)[1] <- "Location"
colnames(FG)[2] <- "Location_Name"
colnames(FG)[3] <- "Mfg_Location_campus"
colnames(FG)[4] <- "Item"
colnames(FG)[5] <- "Description"
colnames(FG)[6] <- "Inventory_Status_Code"
colnames(FG)[7] <- "Hold_Status"
colnames(FG)[8] <- "Current_Inventory_Balance"



# (Path revision needed) Inventory from MicroStrategy (RM) ----

RM <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/Inventory Report for all locations - 10.19.22.xlsx", 
                 col_names = FALSE,
                 sheet = "RM")


RM[-1:-2, ] -> RM
colnames(RM) <- RM[1, ]
RM[-1, ] -> RM

colnames(RM)[1] <- "Location"
colnames(RM)[2] <- "Location_Name"
colnames(RM)[3] <- "Mfg_Location_campus"
colnames(RM)[4] <- "Item"
colnames(RM)[5] <- "Description"
colnames(RM)[6] <- "Inventory_Status_Code"
colnames(RM)[7] <- "Hold_Status"
colnames(RM)[8] <- "Current_Inventory_Balance"

RM %>% 
  dplyr::mutate(Item = sub("^0+", "", Item)) -> RM


# combine FG, RM

rbind(FG, RM) -> inventory_micro


inventory_micro %>% 
  dplyr::mutate(Item = gsub("-", "", Item),  
                ref = paste0(Location, "_", Item),
                campus_ref = paste0(Mfg_Location_campus, "_", Item)) %>% 
  dplyr::relocate(ref, campus_ref) -> inventory_micro


readr::type_convert(inventory_micro) -> inventory_micro

inventory_micro %>% 
  dplyr::mutate(Current_Inventory_Balance = replace(Current_Inventory_Balance, is.na(Current_Inventory_Balance), 0)) -> inventory_micro

# inventory_micro_pivot

reshape2::dcast(inventory_micro, campus_ref ~ Hold_Status , value.var = "Current_Inventory_Balance", sum) %>% 
  dplyr::rename(ref = campus_ref) %>% 
  dplyr::mutate(comp_ref = ref) -> inventory_micro_pivot

inventory_micro_pivot %>% 
  dplyr::rename(Soft_Hold = "Soft Hold",
                Useable_temp = Useable) %>% 
  dplyr::mutate(Useable = Soft_Hold + Useable_temp) -> inventory_micro_pivot


############################################################################################################################
############################################################### ETL ########################################################
############################################################################################################################


# where used count (per loc)
merge(jde_bom, parent_count_1[, c("comp_ref", "parent_count_1")], by = "comp_ref", all.x = TRUE) %>% 
  dplyr::rename(where_used_count_per_loc = parent_count_1) -> jde_bom


# where used count (all loc)
merge(jde_bom, parent_count_2[, c("Component", "parent_count_2")], by = "Component", all.x = TRUE) %>% 
  dplyr::rename(where_used_count_all_loc = parent_count_2) -> jde_bom


# next 28 days open order
merge(jde_bom, Open_Cust_Ord_Pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::rename(next_28_days_open_order = Y) -> jde_bom


# current_month_fcst (mon_a)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_a_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_a_fcst = replace(Mon_a_fcst, is.na(Mon_a_fcst), 0)) -> jde_bom


# next_month_fcst  (mon_b)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_b_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_b_fcst = replace(Mon_b_fcst, is.na(Mon_b_fcst), 0)) -> jde_bom


# next_month_fcst  (mon_c)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_c_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_c_fcst = replace(Mon_c_fcst, is.na(Mon_c_fcst), 0)) -> jde_bom


# next_month_fcst  (mon_d)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_d_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_d_fcst = replace(Mon_d_fcst, is.na(Mon_d_fcst), 0)) -> jde_bom


# next_month_fcst  (mon_e)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_e_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_e_fcst = replace(Mon_e_fcst, is.na(Mon_e_fcst), 0)) -> jde_bom


# next_month_fcst  (mon_f)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_f_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_f_fcst = replace(Mon_f_fcst, is.na(Mon_f_fcst), 0)) -> jde_bom



# next_month_fcst  (mon_g)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_g_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_g_fcst = replace(Mon_g_fcst, is.na(Mon_g_fcst), 0)) -> jde_bom


# next_month_fcst  (mon_h)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_h_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_h_fcst = replace(Mon_h_fcst, is.na(Mon_h_fcst), 0)) -> jde_bom

# next_month_fcst  (mon_i)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_i_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_i_fcst = replace(Mon_i_fcst, is.na(Mon_i_fcst), 0)) -> jde_bom

# next_month_fcst  (mon_j)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_j_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_j_fcst = replace(Mon_j_fcst, is.na(Mon_j_fcst), 0)) -> jde_bom

# next_month_fcst  (mon_k)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_k_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_k_fcst = replace(Mon_k_fcst, is.na(Mon_k_fcst), 0)) -> jde_bom

# next_month_fcst  (mon_l)
merge(jde_bom, DSX_pivot_1[, c("ref", "Mon_l_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Mon_l_fcst = replace(Mon_l_fcst, is.na(Mon_l_fcst), 0)) -> jde_bom


# mon_a_dep_demand 
jde_bom %>% 
  dplyr::mutate(mon_a_dep_demand = pmax(next_28_days_open_order, Mon_a_fcst) * Quantity_w_Scrap) -> jde_bom

# mon_b_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_b_dep_demand = Mon_b_fcst * Quantity_w_Scrap) -> jde_bom

# mon_c_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_c_dep_demand = Mon_c_fcst * Quantity_w_Scrap) -> jde_bom

# mon_d_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_d_dep_demand = Mon_d_fcst * Quantity_w_Scrap) -> jde_bom

# mon_e_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_e_dep_demand = Mon_e_fcst * Quantity_w_Scrap) -> jde_bom

# mon_f_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_f_dep_demand = Mon_f_fcst * Quantity_w_Scrap) -> jde_bom


# mon_g_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_g_dep_demand = Mon_g_fcst * Quantity_w_Scrap) -> jde_bom


# mon_h_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_h_dep_demand = Mon_h_fcst * Quantity_w_Scrap) -> jde_bom


# mon_i_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_i_dep_demand = Mon_i_fcst * Quantity_w_Scrap) -> jde_bom


# mon_j_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_j_dep_demand = Mon_j_fcst * Quantity_w_Scrap) -> jde_bom


# mon_k_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_k_dep_demand = Mon_k_fcst * Quantity_w_Scrap) -> jde_bom


# mon_l_dep_demand
jde_bom %>% 
  dplyr::mutate(mon_l_dep_demand = Mon_l_fcst * Quantity_w_Scrap) -> jde_bom


# FG on Hand
merge(jde_bom, inventory_micro_pivot[, c("ref", "Useable")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Useable = replace(Useable, is.na(Useable), 0)) %>% 
  dplyr::rename(FG_On_Hand = Useable) -> jde_bom


# FG Weeks on Hand
jde_bom %<>% 
  dplyr::mutate(FG_Weeks_On_Hand = FG_On_Hand / ((pmax(next_28_days_open_order, Mon_a_fcst) + Mon_b_fcst + Mon_c_fcst) / 14),
                FG_Weeks_On_Hand = replace(FG_Weeks_On_Hand, is.na(FG_Weeks_On_Hand), 0),
                FG_Weeks_On_Hand = replace(FG_Weeks_On_Hand, is.infinite(FG_Weeks_On_Hand), 0),
                FG_Weeks_On_Hand = replace(FG_Weeks_On_Hand, is.nan(FG_Weeks_On_Hand), 0))


# RM on Hand
merge(jde_bom, inventory_micro_pivot[, c("comp_ref", "Useable")], by = "comp_ref", all.x = TRUE) %>% 
  dplyr::mutate(Useable = replace(Useable, is.na(Useable), 0)) %>% 
  dplyr::rename(RM_On_Hand = Useable) -> jde_bom



# RM_Total_Weeks_On_Hand
# weeks on hand pivot

reshape2::dcast(jde_bom, comp_ref + RM_On_Hand ~ . , value.var = "mon_a_dep_demand", sum) %>% 
  dplyr::rename(sum_month_a_dep_demand = ".") -> mon_a_rm_pivot

reshape2::dcast(jde_bom, comp_ref + RM_On_Hand ~ . , value.var = "mon_b_dep_demand", sum) %>% 
  dplyr::rename(sum_month_b_dep_demand = ".") %>% 
  dplyr::select(-comp_ref, -RM_On_Hand) -> mon_b_rm_pivot

reshape2::dcast(jde_bom, comp_ref + RM_On_Hand ~ . , value.var = "mon_c_dep_demand", sum) %>% 
  dplyr::rename(sum_month_c_dep_demand = ".") %>% 
  dplyr::select(-comp_ref, -RM_On_Hand) -> mon_c_rm_pivot


dplyr::bind_cols(mon_a_rm_pivot, mon_b_rm_pivot, mon_c_rm_pivot) -> weeks_on_hand

weeks_on_hand %>% 
  dplyr::mutate(weeks_on_hand = RM_On_Hand / ((sum_month_a_dep_demand + sum_month_b_dep_demand + sum_month_c_dep_demand) / 14)) %>% 
  dplyr::mutate(weeks_on_hand = replace(weeks_on_hand, is.na(weeks_on_hand), 0),
                weeks_on_hand = replace(weeks_on_hand, is.infinite(weeks_on_hand), 0),
                weeks_on_hand = replace(weeks_on_hand, is.nan(weeks_on_hand), 0)) %>% 
  dplyr::mutate(weeks_on_hand = round(weeks_on_hand, 1)) -> weeks_on_hand



# RM_Total_Weeks_On_Hand
merge(jde_bom, weeks_on_hand[, c("comp_ref", "weeks_on_hand")], by = "comp_ref", all.x = TRUE) %>% 
  dplyr::rename(RM_Total_Weeks_on_Hand = weeks_on_hand) -> jde_bom


######################################################################################################################
############################################## Adding new step 7/26/22 ###############################################
######################################################################################################################

# Adding SKU Status (from exception report) ----

exception_report <- read_excel("C:/Users/lliang/OneDrive - Ventura Foods/R Studio/Source Data/exception report 10.19.22.xlsx")

exception_report[-1:-2, ] -> exception_report
colnames(exception_report) <- exception_report[1, ]
exception_report[-1, -32] -> exception_report

names(exception_report) <- str_replace_all(names(exception_report), c(" " = "_"))

colnames(exception_report)[1] <- "Loc"
colnames(exception_report)[2] <- "Item"

# ref
exception_report %>% 
  dplyr::mutate(ref = paste0(Loc, "_", Item)) %>% 
  dplyr::relocate(ref) -> exception_report

# campus_ref
merge(exception_report, Campus_ref[, c("Loc", "Campus")], by = "Loc", all.x = TRUE) %>% 
  dplyr::mutate(campus_ref = paste0(Campus, "_", Item)) -> exception_report


#### back to jde_bom
# Sku Status
merge(jde_bom, exception_report[, c("ref", "ref")], by = "ref", all.x = TRUE) %>% 
  dplyr::rename(Sku_Status = ref.1) %>% 
  dplyr::mutate(Sku_Status = ifelse(!is.na(Sku_Status), "Active", "Inactive")) -> jde_bom


# Lead Time
exception_report %>% 
  dplyr::mutate(comp_ref = campus_ref) -> exception_report

exception_report[-which(duplicated(exception_report$comp_ref)),] -> exception_report

merge(jde_bom, exception_report[, c("comp_ref", "Leadtime_Days")], by = "comp_ref", all.x = TRUE) %>% 
  dplyr::rename(Lead_time = Leadtime_Days) %>%
  dplyr::mutate(Lead_time = replace(Lead_time, is.na(Lead_time), 0)) -> jde_bom


# category
category_bi %>% 
  dplyr::mutate(Parent_Item_Number = Item) -> category_bi
category_bi[-which(duplicated(category_bi$Parent_Item_Number)),] -> category_bi

merge(jde_bom, category_bi[, c("Parent_Item_Number", "Category")], by = "Parent_Item_Number", all.x = TRUE) -> jde_bom

# Net wt
merge(jde_bom, inventory_model[, c("ref", "Net_wt")], by = "ref", all.x = TRUE) -> jde_bom

# Net wt N/A
# split the data
jde_bom %>% 
  dplyr::filter(!is.na(Net_wt)) -> jde_bom_net_wt_1

jde_bom %>% 
  dplyr::filter(is.na(Net_wt)) -> jde_bom_net_wt_2

merge(jde_bom_net_wt_2, IOM_micro[, c("Parent_Item_Number", "Net_wt")], by = "Parent_Item_Number", all.x = TRUE) -> jde_bom_net_wt_2

jde_bom_net_wt_2 %>% 
  dplyr::select(-Net_wt.x) %>% 
  dplyr::rename(Net_wt = Net_wt.y) -> jde_bom_net_wt_2

rbind(jde_bom_net_wt_1, jde_bom_net_wt_2) -> jde_bom

# Label
jde_bom$Parent_Item_Number -> temp_item
substr(temp_item, nchar(temp_item)-2, nchar(temp_item)) -> temp_item_2
data.frame(temp_item_2) -> temp_item_2
cbind(jde_bom, temp_item_2) -> jde_bom

jde_bom %>% 
  dplyr::rename(Label = temp_item_2) -> jde_bom


# tidy the numbers
jde_bom %>% 
  dplyr::mutate(Unit_Cost = round(Unit_Cost, 2),
                Mon_a_fcst = round(Mon_a_fcst, 2),
                Mon_b_fcst = round(Mon_b_fcst, 2),
                Mon_c_fcst = round(Mon_c_fcst, 2),
                Mon_d_fcst = round(Mon_d_fcst, 2),
                Mon_e_fcst = round(Mon_e_fcst, 2),
                Mon_f_fcst = round(Mon_f_fcst, 2),
                Mon_g_fcst = round(Mon_g_fcst, 2),
                Mon_h_fcst = round(Mon_h_fcst, 2),
                Mon_i_fcst = round(Mon_i_fcst, 2),
                Mon_j_fcst = round(Mon_j_fcst, 2),
                Mon_k_fcst = round(Mon_k_fcst, 2),
                Mon_l_fcst = round(Mon_l_fcst, 2),
                FG_Weeks_On_Hand = round(FG_Weeks_On_Hand, 2)) -> jde_bom

######################################################################################################################
##################################################### final touch ####################################################
######################################################################################################################

jde_bom %>% 
  dplyr::mutate(ref = gsub("_", "-", ref),
                comp_ref = gsub("_", "-", comp_ref)) %>% 
  dplyr::relocate(ref, comp_ref, Sku_Status, Category, Label, where_used_count_per_loc, where_used_count_all_loc, Business_Unit, Level, Parent_Item_Number,
                  Parent_Description, UOM, Net_wt ,FG_On_Hand, FG_Weeks_On_Hand, Component, Component_Description, Commodity_Class,
                  UM, Lead_time, RM_On_Hand, RM_Total_Weeks_on_Hand, Stocking_Type, Percent_Scrap, Quantity_Per, Quantity_w_Scrap, Unit_Cost,
                  next_28_days_open_order, Mon_a_fcst, Mon_b_fcst, Mon_c_fcst, Mon_d_fcst, Mon_e_fcst, Mon_f_fcst, Mon_g_fcst, Mon_h_fcst,
                  Mon_i_fcst, Mon_j_fcst, Mon_k_fcst, Mon_l_fcst,
                  mon_a_dep_demand, mon_b_dep_demand, mon_c_dep_demand, mon_d_dep_demand, mon_e_dep_demand, mon_f_dep_demand,
                  mon_g_dep_demand, mon_h_dep_demand, mon_i_dep_demand, mon_j_dep_demand, mon_k_dep_demand, mon_l_dep_demand) %>% 
  dplyr::mutate(FG_Weeks_On_Hand = round(FG_Weeks_On_Hand, 1),
                Mon_a_fcst = round(Mon_a_fcst, 0),
                Mon_b_fcst = round(Mon_b_fcst, 0),
                Mon_c_fcst = round(Mon_c_fcst, 0),
                Mon_d_fcst = round(Mon_d_fcst, 0),
                Mon_e_fcst = round(Mon_e_fcst, 0),
                Mon_f_fcst = round(Mon_f_fcst, 0),
                Mon_g_fcst = round(Mon_g_fcst, 0),
                Mon_h_fcst = round(Mon_h_fcst, 0),
                Mon_i_fcst = round(Mon_i_fcst, 0),
                Mon_j_fcst = round(Mon_j_fcst, 0),
                Mon_k_fcst = round(Mon_k_fcst, 0),
                Mon_l_fcst = round(Mon_l_fcst, 0)) %>% 
  dplyr::mutate(Component = as.integer(Component)) -> jde_bom





colnames(jde_bom)[1]<-"ref"
colnames(jde_bom)[2]<-"comp ref"
colnames(jde_bom)[3]<-"SKU Status"
colnames(jde_bom)[4]<-"Category"
colnames(jde_bom)[5]<-"Label"
colnames(jde_bom)[6]<-"where used count (per loc)"
colnames(jde_bom)[7]<-"where used count (all loc)"
colnames(jde_bom)[8]<-"Business Unit"
colnames(jde_bom)[9]<-"Level"
colnames(jde_bom)[10]<-"Parent Item Number"
colnames(jde_bom)[11]<-"Parent Description"
colnames(jde_bom)[12]<-"UOM"
colnames(jde_bom)[13]<-"Net_wt"
colnames(jde_bom)[14]<-"FG On Hand"
colnames(jde_bom)[15]<-"FG Weeks on Hand"
colnames(jde_bom)[16]<-"Component"
colnames(jde_bom)[17]<-"Component Description"
colnames(jde_bom)[18]<-"Commodity Class"
colnames(jde_bom)[19]<-"UM"
colnames(jde_bom)[20]<-"Lead Time"
colnames(jde_bom)[21]<-"RM On Hand"
colnames(jde_bom)[22]<-"RM Total Weeks on Hand"
colnames(jde_bom)[23]<-"Stocking Type"
colnames(jde_bom)[24]<-"Percent Scrap"
colnames(jde_bom)[25]<-"Quantity Per"
colnames(jde_bom)[26]<-"Quantity w/ Scrap"
colnames(jde_bom)[27]<-"Unit Cost"
colnames(jde_bom)[28]<-"next 28 days open order"
colnames(jde_bom)[29]<-"mon_a fcst"
colnames(jde_bom)[30]<-"mon_b fcst"
colnames(jde_bom)[31]<-"mon_c fcst"
colnames(jde_bom)[32]<-"mon_d fcst"
colnames(jde_bom)[33]<-"mon_e fcst"
colnames(jde_bom)[34]<-"mon_f fcst"
colnames(jde_bom)[35]<-"mon_g fcst"
colnames(jde_bom)[36]<-"mon_h fcst"
colnames(jde_bom)[37]<-"mon_i fcst"
colnames(jde_bom)[38]<-"mon_j fcst"
colnames(jde_bom)[39]<-"mon_k fcst"
colnames(jde_bom)[40]<-"mon_l fcst"
colnames(jde_bom)[41]<-"mon_a dep demand"
colnames(jde_bom)[42]<-"mon_b dep demand"
colnames(jde_bom)[43]<-"mon_c dep demand"
colnames(jde_bom)[44]<-"mon_d dep demand"
colnames(jde_bom)[45]<-"mon_e dep demand"
colnames(jde_bom)[46]<-"mon_f dep demand"
colnames(jde_bom)[47]<-"mon_g dep demand"
colnames(jde_bom)[48]<-"mon_h dep demand"
colnames(jde_bom)[49]<-"mon_i dep demand"
colnames(jde_bom)[50]<-"mon_j dep demand"
colnames(jde_bom)[51]<-"mon_k dep demand"
colnames(jde_bom)[52]<-"mon_l dep demand"


writexl::write_xlsx(jde_bom, "Bill of Material.xlsx")