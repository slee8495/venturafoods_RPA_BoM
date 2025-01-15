library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(officer)
library(openxlsx)
library(lubridate)
library(magrittr)
library(skimr)
library(rio)

specific_date <- as.Date("2025-01-14")

##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################

## Supplier Address Book 
supplier_address  <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Address Book/Address Book - 2025.01.07.xlsx",
                                sheet = "supplier")

## FG_ref_to_mpg_ref 

complete_sku_list <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01142025/Complete SKU list - Linda.xlsx")
complete_sku_list[-1, ] -> complete_sku_list
colnames(complete_sku_list) <- complete_sku_list[1, ]
complete_sku_list[-1, ] -> complete_sku_list

complete_sku_list %>% 
  janitor::clean_names() %>% 
  dplyr::slice(-1) %>% 
  dplyr::select(item_location_no_v2, product_label_sku, product_manufacturing_location) %>%
  dplyr::rename(loc = item_location_no_v2,
                sku = product_label_sku,
                mfg_loc = product_manufacturing_location) %>% 
  dplyr::mutate(sku = gsub("-", "", sku)) %>% 
  dplyr::mutate(ref = paste0(loc, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku)) -> FG_ref_to_mfg_ref




# Campus_ref pulling 

Campus_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Campus reference.xlsx") %>% 
  readr::type_convert()



colnames(Campus_ref)[1] <- "Loc"
colnames(Campus_ref)[2] <- "Description"
colnames(Campus_ref)[3] <- "Campus"

# (Path revision needed) Category (From BI) ---- 
category_bi <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2022/12.19.2022/BI Category and Platform and pack size.xlsx")

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

# S:Drive - Supply Chain Project - Logistics - SCP - Cost Saving Reporting 


inventory_model_data <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx",
                                   col_names = FALSE, sheet = "Fin Goods")


inventory_model_data[-1:-7, ] -> inventory_model_data
colnames(inventory_model_data) <- inventory_model_data[1, ]
inventory_model_data[-1, ] -> inventory_model_data

inventory_model_data %>% 
  dplyr::select("Ship Ref", "Net Wt") %>% 
  dplyr::rename(ref = "Ship Ref",
                Net_wt = "Net Wt") %>% 
  dplyr::mutate(ref = gsub("-", "_", ref),
                Net_wt = as.numeric(Net_wt)) -> inventory_model

# (Path revision needed) IOM MicroStrategy ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/07915A52DE47AA1CDB4AB082191E4EBA/K271--K264
IOM_micro <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01142025/IOM Data Extract.xlsx")

IOM_micro[-1, ] -> IOM_micro
colnames(IOM_micro) <- IOM_micro[1, ]
IOM_micro[-1, ] -> IOM_micro

IOM_micro %>% 
  dplyr::select("Product Label (SKU)", "FG Net Weight") %>% 
  dplyr::rename(Parent_Item_Number = "Product Label (SKU)",
                Net_wt = "FG Net Weight") %>% 
  dplyr::mutate(Net_wt = as.numeric(Net_wt),
                Parent_Item_Number = gsub("-", "", Parent_Item_Number)) -> IOM_micro


# Exception Report ----

exception_report <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE Exception report extract/2025/exception report 2025.01.07.xlsx")

exception_report[-1:-2, ] -> exception_report
colnames(exception_report) <- exception_report[1, ]
exception_report[-1, -32] -> exception_report

names(exception_report) <- str_replace_all(names(exception_report), c(" " = "_"))

colnames(exception_report)[1] <- "Loc"
colnames(exception_report)[2] <- "Item"


exception_report -> exception_report_supplier
exception_report -> exception_report_lead_time


# (Path revision needed) DSX Forecast backup ----

DSX_Forecast_Backup <- read_excel(
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2025/DSX Forecast Backup - 2025.01.13.xlsx")

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


# Filter to start from Last month
last_month_start <- as.Date(format(Sys.Date(), "%Y-%m-01")) - months(1)

# Convert Forecast_Month_Year_Code_Segment_ID to a date format for filtering
DSX_Forecast_Backup %>%
  dplyr::mutate(Forecast_Date = as.Date(paste0(substr(Forecast_Month_Year_Code_Segment_ID, 1, 4), "-", substr(Forecast_Month_Year_Code_Segment_ID, 5, 6), "-01"))) %>% 
  dplyr::filter(Forecast_Date >= last_month_start) -> DSX_Forecast_Backup





# DSX Pivot table
reshape2::dcast(DSX_Forecast_Backup, mfg_ref ~ Forecast_Month_Year_Code_Segment_ID , value.var = "Adjusted_Forecast_Cases", sum) -> DSX_pivot_1


last_month <- floor_date(specific_date %m-% months(1), "month")
last_month_year <- year(last_month)
last_month_month <- month(last_month)
# Format as YYYYMM
start_column <- sprintf("%d%02d", last_month_year, last_month_month)

DSX_pivot_1 %>%
  select(mfg_ref, starts_with(start_column):last_col()) -> DSX_pivot_1


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


# https://edgeanalytics.venturafoods.com:443/MicroStrategy/servlet/mstrWeb?evt=4058&src=mstrWeb.4058&_subscriptionID=1ADEEE1E6046707D2EE259B1A3D4F767&reportViewMode=1&Server=ENV-323771LAIO1USE2&Project=VF%20Intelligent%20Enterprise&Port=39321&share=1
# (Path revision needed) Opencustord ----
Open_Cust_Ord <- read.xlsx("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01142025/US and CAN OO BT where status _ J.xlsx",
                           colNames = FALSE)


Open_Cust_Ord %>% 
  dplyr::slice(c(-1, -3)) -> Open_Cust_Ord

colnames(Open_Cust_Ord) <- Open_Cust_Ord[1, ]
Open_Cust_Ord[-1, ] -> Open_Cust_Ord

Open_Cust_Ord %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
  dplyr::mutate(ref = paste0(location, "_", product_label_sku)) %>% 
  dplyr::mutate(oo_cases = as.double(oo_cases),
                oo_cases = ifelse(is.na(oo_cases), 0, oo_cases),
                b_t_open_order_cases = as.double(b_t_open_order_cases),
                b_t_open_order_cases = ifelse(is.na(b_t_open_order_cases), 0, b_t_open_order_cases)) %>%
  dplyr::mutate(Qty = oo_cases + b_t_open_order_cases) %>% 
  dplyr::mutate(sales_order_requested_ship_date = as_date(as.integer(sales_order_requested_ship_date), origin = "1899-12-30")) %>% 
  dplyr::mutate(year = year(sales_order_requested_ship_date),
                month = month(sales_order_requested_ship_date),
                day = day(sales_order_requested_ship_date),
                month_year = paste0(month, "_", year)) %>% 
  dplyr::rename(ProductSkuCode = product_label_sku,
                Loc = location,
                date = sales_order_requested_ship_date) %>% 
  dplyr::select(ref, ProductSkuCode, Loc, Qty, date, year, month, day, month_year) %>% 
  tibble::as_tibble() %>% 
  dplyr::mutate(year = as.character(year),
                month = as.character(month),
                day = as.character(day)) -> Open_Cust_Ord

Open_Cust_Ord %>% 
  dplyr::group_by(ref, ProductSkuCode, Loc, date, year, month, day, month_year) %>% 
  dplyr::summarise(Qty = sum(Qty)) %>% 
  dplyr::relocate(Qty, .after = "Loc") -> Open_Cust_Ord



# (Path revision needed) Sales and Open orders cube from Micro (Canada only) ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/46031E5A134A6DD24564938529CF0EB8
canada_micro <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01142025/Canada Open Orders.xlsx", 
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
  dplyr::mutate(next_28_days = ifelse(date <= specific_date + 28, "Y", "N")) -> Open_Cust_Ord

reshape2::dcast(Open_Cust_Ord, ref ~ next_28_days, value.var = "Qty", sum) -> Open_Cust_Ord_Pivot



# (Path revision needed) Read JDE BoM ----

jde_bom_us <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01142025/jde_us.xlsx", 
                         col_names = FALSE)


jde_bom_us[-1:-2, ] -> jde_bom_us
colnames(jde_bom_us) <- jde_bom_us[1, ] 
jde_bom_us[-1, ] -> jde_bom_us
jde_bom_us[ , c(-4,-16)] -> jde_bom_us

names(jde_bom_us) <- stringr::str_replace_all(names(jde_bom_us), c(" " = "_"))
type_convert(jde_bom_us) -> jde_bom_us


colnames(jde_bom_us)[13] <- "Quantity_w_Scrap"

jde_bom_us %>%
  mutate(
    needs_correction = is.na(Unit_Cost),
    Stocking_Type = ifelse(needs_correction, jde_bom_us$Commodity_Class, Stocking_Type),
    Percent_Scrap = ifelse(needs_correction, jde_bom_us$UM, Percent_Scrap),
    Quantity_Per = ifelse(needs_correction, jde_bom_us$Stocking_Type, Quantity_Per),
    Quantity_w_Scrap = ifelse(needs_correction, jde_bom_us$Percent_Scrap, Quantity_w_Scrap),
    Unit_Cost = ifelse(needs_correction, jde_bom_us$Quantity_Per, Unit_Cost)
  ) %>%
  select(-needs_correction) -> jde_bom_us


jde_bom_us %>% 
  dplyr::mutate(ref = paste0(Business_Unit, "_", Parent_Item_Number),
                comp_ref = paste0(Business_Unit, "_", Component)) -> jde_bom_us







jde_bom_canada <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01142025/jde_canada.xlsx", 
                             col_names = FALSE)


jde_bom_canada[-1:-2, ] -> jde_bom_canada
colnames(jde_bom_canada) <- jde_bom_canada[1, ] 
jde_bom_canada[-1, ] -> jde_bom_canada
jde_bom_canada[ , c(-4,-16)] -> jde_bom_canada

names(jde_bom_canada) <- stringr::str_replace_all(names(jde_bom_canada), c(" " = "_"))
type_convert(jde_bom_canada) -> jde_bom_canada



colnames(jde_bom_canada)[13] <- "Quantity_w_Scrap"


jde_bom_canada %>%
  mutate(
    needs_correction = is.na(Unit_Cost),
    Stocking_Type = ifelse(needs_correction, jde_bom_canada$Commodity_Class, Stocking_Type),
    Percent_Scrap = ifelse(needs_correction, jde_bom_canada$UM, Percent_Scrap),
    Quantity_Per = ifelse(needs_correction, jde_bom_canada$Stocking_Type, Quantity_Per),
    Quantity_w_Scrap = ifelse(needs_correction, jde_bom_canada$Percent_Scrap, Quantity_w_Scrap),
    Unit_Cost = ifelse(needs_correction, jde_bom_canada$Quantity_Per, Unit_Cost)
  ) %>%
  select(-needs_correction) -> jde_bom_canada

jde_bom_canada %>% 
  dplyr::mutate(ref = paste0(Business_Unit, "_", Parent_Item_Number),
                comp_ref = paste0(Business_Unit, "_", Component)) -> jde_bom_canada


rbind(jde_bom_us, jde_bom_canada) -> jde_bom




# parent count
jde_bom %>% 
  dplyr::count(comp_ref, Parent_Item_Number) %>% 
  dplyr::group_by(comp_ref) %>%
  dplyr::summarize(parent_count_1 = n_distinct(Parent_Item_Number)) -> parent_count_1

jde_bom %>% 
  dplyr::count(Component, Parent_Item_Number) %>% 
  dplyr::group_by(Component) %>%
  dplyr::summarize(parent_count_2 = n_distinct(Parent_Item_Number)) -> parent_count_2



# Inventory Status Code table
Inventory_Status_Code <- c("", "Q", "W")
Hold_Status <- c("Useable", "Hard Hold", "Soft Hold")

data.frame(Inventory_Status_Code, Hold_Status) -> inventory_status_table



inventory_micro_rm <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2025.01.14.xlsx",
                                 sheet = "RM")






inventory_micro_rm[-1, ] -> inventory_micro_rm
colnames(inventory_micro_rm) <- inventory_micro_rm[1, ]
inventory_micro_rm[-1, ] -> inventory_micro_rm




inventory_micro_rm %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>% 
  dplyr::mutate(item = str_replace(item, "^0+(?!$)", "")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item),
                campus_ref = paste0(campus_no, "_", item)) %>% 
  dplyr::select(location, item, description, campus_no, inventory_hold_status, current_inventory_balance, ref, campus_ref) %>% 
  dplyr::mutate(current_inventory_balance = as.numeric(current_inventory_balance)) -> inventory_micro


inventory_micro %>% 
  dplyr::rename(Location = location, 
                Item = item,
                Description = description,
                Mfg_Location_campus = campus_no,
                Hold_Status = inventory_hold_status,
                Current_Inventory_Balance = current_inventory_balance) -> inventory_micro


inventory_micro %>% 
  filter(!str_starts(Description, "PWS ") & 
           !str_starts(Description, "SUB ") & 
           !str_starts(Description, "THW ") & 
           !str_starts(Description, "PALLET")) -> inventory_micro





reshape2::dcast(inventory_micro, campus_ref ~ Hold_Status , value.var = "Current_Inventory_Balance", sum) %>%
  dplyr::rename(ref = campus_ref) %>%
  dplyr::mutate(comp_ref = ref) -> inventory_micro_pivot

inventory_micro_pivot %>%
  dplyr::rename(Soft_Hold = "Soft Hold",
                Hard_Hold = "Hard Hold",
                Useable_temp = Useable) %>%
  dplyr::mutate(Useable = Soft_Hold + Useable_temp) -> inventory_micro_pivot




inventory_micro_fg <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2025.01.14.xlsx",
                                 sheet = "FG")





inventory_micro_fg[-1, ] -> inventory_micro_fg
colnames(inventory_micro_fg) <- inventory_micro_fg[1, ]
inventory_micro_fg[-1, ] -> inventory_micro_fg




inventory_micro_fg %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>% 
  dplyr::mutate(ref = paste0(location, "_", item),
                campus_ref = paste0(campus_no, "_", item)) %>% 
  dplyr::select(location, item, description, campus_no, inventory_hold_status, current_inventory_balance, ref, campus_ref) %>% 
  dplyr::mutate(current_inventory_balance = as.numeric(current_inventory_balance)) -> inventory_micro_fg


inventory_micro_fg %>% 
  dplyr::rename(Location = location, 
                Item = item,
                Description = description,
                Mfg_Location_campus = campus_no,
                Hold_Status = inventory_hold_status,
                Current_Inventory_Balance = current_inventory_balance) -> inventory_micro_fg


inventory_micro_fg %>% 
  filter(!str_starts(Description, "PWS ") & 
           !str_starts(Description, "SUB ") & 
           !str_starts(Description, "THW ") & 
           !str_starts(Description, "PALLET")) -> inventory_micro_fg





reshape2::dcast(inventory_micro_fg, campus_ref ~ Hold_Status , value.var = "Current_Inventory_Balance", sum) %>%
  dplyr::rename(ref = campus_ref) %>%
  dplyr::mutate(comp_ref = ref) -> inventory_micro_fg_pivot

inventory_micro_fg_pivot %>%
  dplyr::rename(Soft_Hold = "Soft Hold",
                Hard_Hold = "Hard Hold",
                Useable_temp = Useable) %>%
  dplyr::mutate(Useable = Soft_Hold + Useable_temp) -> inventory_micro_fg_pivot


################# inv_bal for 25, 55 label ###############


lot_status_code <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Lot Status Code.xlsx")

lot_status_code %>% 
  janitor::clean_names() %>% 
  dplyr::select(lot_status, hard_soft_hold) %>% 
  dplyr::mutate(lot_status = ifelse(is.na(lot_status), "Useable", lot_status),
                hard_soft_hold = ifelse(is.na(hard_soft_hold), "Useable", hard_soft_hold)) %>% 
  dplyr::rename(status = lot_status) -> lot_status_code



jde_inv_for_25_55_label <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/JDE Inventory Lot Detail - 2025.01.14.xlsx")

jde_inv_for_25_55_label[-1:-5, ] -> jde_inv_for_25_55_label
colnames(jde_inv_for_25_55_label) <- jde_inv_for_25_55_label[1, ]
jde_inv_for_25_55_label[-1, ] -> jde_inv_for_25_55_label


jde_inv_for_25_55_label %>% 
  janitor::clean_names() %>% 
  dplyr::rename(b_p = bp,
                item = item_number) %>% 
  dplyr::mutate(status = ifelse(is.na(status), "Useable", status)) %>% 
  dplyr::mutate(item = as.numeric(item),
                on_hand = as.numeric(on_hand),
                b_p = as.numeric(b_p)) %>% 
  dplyr::filter(!is.na(item)) %>% 
  dplyr::left_join(lot_status_code, by = "status") %>% 
  dplyr::select(-status) %>% 
  pivot_wider(names_from = hard_soft_hold, values_from = on_hand, values_fn = list(on_hand = sum)) %>% 
  janitor::clean_names() %>% 
  replace_na(list(useable = 0, soft_hold = 0, hard_hold = 0)) %>% 
  dplyr::left_join(exception_report %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(item, mpf_or_line) %>% 
                     dplyr::rename(label = mpf_or_line) %>% 
                     dplyr::mutate(item = as.double(item)) %>% 
                     dplyr::filter(label == "LBL") %>% 
                     dplyr::distinct(item, label)) %>% 
  dplyr::filter(!is.na(label)) %>% 
  dplyr::select(-label) %>% 
  dplyr::mutate(ref = paste0(b_p, "_", item)) %>% 
  dplyr::mutate(useable = useable + soft_hold) %>% 
  dplyr::mutate(on_hand = useable + hard_hold) %>%
  dplyr::select(ref, hard_hold, soft_hold, useable) %>% 
  dplyr::rename(Hard_Hold = hard_hold,
                Soft_Hold = soft_hold,
                Useable = useable) %>% 
  dplyr::mutate(Useable_temp = Useable,
                comp_ref = ref) %>% 
  dplyr::relocate(ref, Hard_Hold, Soft_Hold, Useable_temp, comp_ref, Useable) -> label_25_55_pivot


rbind(inventory_micro_pivot, label_25_55_pivot) -> inventory_micro_pivot



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
merge(jde_bom, inventory_micro_fg_pivot[, c("ref", "Useable")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Useable = replace(Useable, is.na(Useable), 0)) %>% 
  dplyr::rename(FG_On_Hand = Useable) -> jde_bom


# FG Weeks on Hand
jde_bom %>% 
  dplyr::mutate(FG_Weeks_On_Hand = FG_On_Hand / ((pmax(next_28_days_open_order, Mon_a_fcst) + Mon_b_fcst + Mon_c_fcst) / 14),
                FG_Weeks_On_Hand = replace(FG_Weeks_On_Hand, is.na(FG_Weeks_On_Hand), 0),
                FG_Weeks_On_Hand = replace(FG_Weeks_On_Hand, is.infinite(FG_Weeks_On_Hand), 0),
                FG_Weeks_On_Hand = replace(FG_Weeks_On_Hand, is.nan(FG_Weeks_On_Hand), 0)) -> jde_bom


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



# ref
exception_report %>% 
  dplyr::mutate(ref = paste0(Loc, "_", Item)) %>% 
  dplyr::relocate(ref) -> exception_report

exception_report[!duplicated(exception_report[,c("ref")]),] -> exception_report

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
  dplyr::mutate(Unit_Cost = round(Unit_Cost, 4),
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


# Category & Platform
completed_sku_list <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01142025/Complete SKU list - Linda.xlsx")
completed_sku_list[-1:-2, ]  %>% 
  janitor::clean_names() %>% 
  dplyr::select(x6, x9, x11) %>% 
  dplyr::rename(Parent_Item_Number = x6,
                Category = x9,
                Platform = x11) %>% 
  dplyr::mutate(Parent_Item_Number = gsub("-", "", Parent_Item_Number)) -> completed_sku_list

completed_sku_list[!duplicated(completed_sku_list[,c("Parent_Item_Number")]),] -> completed_sku_list

completed_sku_list %>% 
  dplyr::select(Parent_Item_Number, Category) -> completed_sku_list_category


completed_sku_list %>% 
  dplyr::select(Parent_Item_Number, Platform) -> completed_sku_list_platform


jde_bom %>% 
  dplyr::select(-Category) %>% 
  dplyr::left_join(completed_sku_list_category) %>% 
  dplyr::left_join(completed_sku_list_platform) -> jde_bom


# Net wt
inventory_model_data %>% 
  janitor::clean_names() %>% 
  dplyr::select(ship_ref, net_wt) %>% 
  dplyr::mutate(ship_ref = gsub("-", "_", ship_ref)) %>% 
  dplyr::rename(ref = ship_ref) %>% 
  dplyr::mutate(net_wt = gsub(";", "", net_wt)) -> inventory_model_net_wt

inventory_model_net_wt[!duplicated(inventory_model_net_wt[,c("ref")]),] -> inventory_model_net_wt

jde_bom %>% 
  dplyr::left_join(inventory_model_net_wt) %>% 
  dplyr::mutate(Net_wt = ifelse(is.na(Net_wt), net_wt, Net_wt)) %>% 
  dplyr::select(-net_wt) -> jde_bom


######################################################################################################################
##################################################### update 5/24/23 #################################################
######################################################################################################################


# Net Wt code update
pre_bom <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01072025/Bill of Material_01072025.xlsx")

pre_bom %>% 
  data.frame() %>% 
  dplyr::select(Parent.Item.Number, Net_wt) %>% 
  dplyr::rename(Parent_Item_Number = Parent.Item.Number,
                Net_wt_2 = Net_wt) -> pre_bom_net_wt

pre_bom_net_wt[!duplicated(pre_bom_net_wt[,c("Parent_Item_Number")]),] -> pre_bom_net_wt

jde_bom %>% 
  dplyr::left_join(pre_bom_net_wt) %>% 
  dplyr::mutate(Net_wt = ifelse(is.na(Net_wt), Net_wt_2, Net_wt)) %>% 
  dplyr::select(-Net_wt_2) -> jde_bom



# - Supplier
exception_report %>% 
  dplyr::mutate(Loc = as.double(Loc)) %>%
  dplyr::select(Loc, Campus, campus_ref, Supplier) %>% 
  dplyr::arrange(desc(Loc)) %>% 
  dplyr::mutate(loc_supplier = paste0(Loc, "_", Supplier)) %>% 
  dplyr::select(campus_ref, Supplier) %>% 
  dplyr::rename(comp_ref = campus_ref) -> exception_report_2


jde_bom %>% 
  dplyr::left_join(exception_report_2) %>% 
  dplyr::mutate(Supplier = ifelse(is.na(Supplier), "NA", Supplier)) -> jde_bom




######################################################################################################################
##################################################### update 5/31/23 #################################################
######################################################################################################################


pre_bom %>% 
  data.frame() %>% 
  dplyr::select(Parent.Item.Number, Category, Platform) %>% 
  dplyr::rename(Parent_Item_Number = Parent.Item.Number,
                category = Category,
                platform = Platform) -> pre_bom_category_platform

pre_bom_category_platform[!duplicated(pre_bom_category_platform[,c("Parent_Item_Number")]),] -> pre_bom_category_platform

jde_bom %>% 
  dplyr::left_join(pre_bom_category_platform) %>% 
  dplyr::mutate(Category = ifelse(is.na(Category), category, Category),
                Platform = ifelse(is.na(Platform), platform, Platform)) %>% 
  dplyr::select(-category, -platform) -> jde_bom


######################################################################################################################
##################################################### update 8/16/23 #################################################
######################################################################################################################
supplier_address %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 2) %>% 
  dplyr::rename(Supplier = address_number,
                supplier_name = alpha_name) %>% 
  dplyr::mutate(Supplier = as.character(Supplier)) -> supplier_name

jde_bom %>% 
  dplyr::left_join(supplier_name) %>% 
  dplyr::mutate(supplier_name = ifelse(is.na(supplier_name), "NA", supplier_name)) -> jde_bom


######################################################################################################################
##################################################### update 8/24/23 #################################################
######################################################################################################################



jde_bom %>% 
  dplyr::mutate(Component = sub("^0+", "", Component)) %>% 
  dplyr::mutate(comp_ref = paste0(Business_Unit, "_", Component)) -> jde_bom


###################################### Lead Time 04/05/2024 #################################


exception_report_lead_time %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(ref = paste0(loc, "_", item)) %>% 
  dplyr::left_join(Campus_ref %>% select(Loc, Campus), by = c("loc" = "Loc")) %>% 
  dplyr::mutate(campus_ref = paste0(Campus, "_", item)) %>% 
  dplyr::mutate(leadtime_days = ifelse(is.na(leadtime_days), 0, as.numeric(leadtime_days))) %>% 
  dplyr::arrange(desc(leadtime_days)) %>% 
  dplyr::select(campus_ref, leadtime_days) %>% 
  dplyr::group_by(campus_ref) %>%
  dplyr::slice_max(leadtime_days, n = 1) %>%
  dplyr::ungroup() %>% 
  dplyr::distinct(campus_ref, .keep_all = TRUE) -> exception_report_lead_time


jde_bom %>% 
  dplyr::select(-Lead_time) %>% 
  dplyr::left_join(exception_report_lead_time %>% rename(Lead_time = leadtime_days), by = c("comp_ref" = "campus_ref")) %>% 
  dplyr::mutate(Lead_time = ifelse(is.na(Lead_time), 0, Lead_time)) -> jde_bom



################  Spplier 7/29/2024 ##################
exception_report_supplier %>% 
  janitor::clean_names() %>% 
  dplyr::select(loc, item, supplier, mpf_or_line) %>% 
  dplyr::mutate(comp_ref = paste0(loc, "_", item)) %>% 
  dplyr::distinct(comp_ref, .keep_all = TRUE) -> exception_report_supplier



## -Inventory logic

jde_bom %>%
  dplyr::select(-Supplier, -supplier_name) %>%
  dplyr::left_join(exception_report_supplier %>% dplyr::select(comp_ref, mpf_or_line), by = c("comp_ref" = "comp_ref")) %>%
  dplyr::left_join(exception_report_supplier %>% dplyr::select(comp_ref, supplier), by = c("comp_ref" = "comp_ref")) %>%
  dplyr::left_join(supplier_name, by = c("supplier" = "Supplier")) %>%
  dplyr::mutate(ab = ifelse(str_detect(supplier_name, regex("-INVENTORY", ignore_case = TRUE)), "a", "b")) %>%
  dplyr::mutate(ab = ifelse(is.na(supplier),  "b", ab)) %>% 
  dplyr::mutate(mpf_ref = paste0(mpf_or_line, "_", Component)) -> jde_bom_sup_1

# Separate data frames based on the condition
jde_bom_sup_1 %>%
  dplyr::filter(ab == "a") -> jde_bom_sup_1_a

jde_bom_sup_1 %>%
  dplyr::filter(ab == "b") -> jde_bom_sup_1_b

# Perform the additional join for 'a' subset
jde_bom_sup_1_a %>%
  dplyr::select(-supplier, -supplier_name) %>% 
  dplyr::left_join(exception_report_supplier %>%
                     dplyr::select(comp_ref, supplier) %>% 
                     dplyr::distinct(comp_ref, .keep_all = TRUE),
                   by = c("mpf_ref" = "comp_ref")) %>% 
  dplyr::left_join(supplier_name, by = c("supplier" = "Supplier")) %>%
  dplyr::relocate(c(supplier, supplier_name, ab, mpf_ref), .after = mpf_or_line) -> jde_bom_sup_1_a

rbind(jde_bom_sup_1_a, jde_bom_sup_1_b) -> jde_bom_sup_1


## Second try: "-inventory"

jde_bom_sup_1 %>% 
  dplyr::mutate(ab = ifelse(str_detect(supplier_name, regex("-INVENTORY", ignore_case = TRUE)), "a", "b")) %>%
  dplyr::mutate(ab = ifelse(is.na(supplier),  "b", ab)) -> jde_bom_sup_2


jde_bom_sup_2 %>%
  dplyr::filter(ab == "a") -> jde_bom_sup_2_a

jde_bom_sup_2 %>%
  dplyr::filter(ab == "b") -> jde_bom_sup_2_b


jde_bom_sup_2_a %>% 
  dplyr::select(Component) %>% 
  dplyr::distinct()  %>% 
  dplyr::pull(Component) -> second_try_inventory




# Function to find destination supplier
find_destination_supplier <- function(comp_ref, exception_report_supplier) {
  # Split the comp_ref into loc and item
  split_ref <- str_split(comp_ref, "_", simplify = TRUE)
  current_loc <- split_ref[1]
  item <- split_ref[2]
  
  visited_locs <- c()
  
  cat("Processing comp_ref:", comp_ref, "/n")
  
  while (TRUE) {
    if (current_loc %in% visited_locs) {
      cat("Cycle detected at loc:", current_loc, "/n")
      return(NA)
    }
    
    visited_locs <- c(visited_locs, current_loc)
    
    current_row <- exception_report_supplier %>%
      filter(loc == current_loc, item == !!item)
    
    if (nrow(current_row) == 0) {
      cat("No row found for loc:", current_loc, "and item:", item, "/n")
      return(NA)
    }
    
    mpf_or_line <- current_row$mpf_or_line
    supplier <- current_row$supplier
    
    cat("Current loc:", current_loc, "mpf_or_line:", mpf_or_line, "supplier:", supplier, "/n")
    
    # Check if mpf_or_line is NA or non-numeric
    if (is.na(mpf_or_line) || !str_detect(mpf_or_line, "^[0-9]+$")) {
      return(supplier)
    }
    
    current_loc <- mpf_or_line
  }
}



exception_report_supplier %>%
  filter(item %in% second_try_inventory) %>%
  pull(comp_ref) -> comp_refs

supplier_destination_result <- sapply(comp_refs, find_destination_supplier, exception_report_supplier = exception_report_supplier)

supplier_destination_df <- data.frame(
  comp_ref = names(supplier_destination_result),
  supplier = unname(supplier_destination_result)
)

# Clean column names
supplier_destination_df %>% 
  janitor::clean_names() -> supplier_destination_df




jde_bom_sup_2_a %>% 
  dplyr::select(-supplier, -supplier_name) %>% 
  dplyr::left_join(supplier_destination_df, by = c("mpf_ref" = "comp_ref")) %>% 
  dplyr::left_join(supplier_name, by = c("supplier" = "Supplier")) %>%
  dplyr::relocate(c(supplier, supplier_name, ab, mpf_ref), .after = mpf_or_line) -> jde_bom_sup_2_a



rbind(jde_bom_sup_2_a, jde_bom_sup_2_b) -> jde_bom


## vf logic

jde_bom %>%
  dplyr::mutate(ab = ifelse(str_detect(supplier_name, regex("vf", ignore_case = TRUE)), "a", "b")) %>%
  dplyr::mutate(ab = ifelse(is.na(supplier),  "b", ab))-> jde_bom_sup_3


jde_bom_sup_3 %>%
  dplyr::filter(ab == "a") -> jde_bom_sup_3_a

jde_bom_sup_3 %>%
  dplyr::filter(ab == "b") -> jde_bom_sup_3_b


jde_bom_sup_3_a %>%
  dplyr::select(-supplier, -supplier_name) %>% 
  dplyr::left_join(exception_report_supplier %>%
                     dplyr::select(comp_ref, supplier) %>% 
                     dplyr::distinct(comp_ref, .keep_all = TRUE),
                   by = c("mpf_ref" = "comp_ref")) %>% 
  dplyr::left_join(supplier_name, by = c("supplier" = "Supplier")) %>%
  dplyr::relocate(c(supplier, supplier_name, ab, mpf_ref), .after = mpf_or_line) -> jde_bom_sup_3_a

rbind(jde_bom_sup_3_a, jde_bom_sup_3_b) -> jde_bom_sup_3





## Second try: "vf"

jde_bom_sup_3 %>% 
  dplyr::mutate(ab = ifelse(str_detect(supplier_name, regex("vf", ignore_case = TRUE)), "a", "b")) %>%
  dplyr::mutate(ab = ifelse(is.na(supplier),  "b", ab)) -> jde_bom_sup_3


jde_bom_sup_3 %>%
  dplyr::filter(ab == "a") -> jde_bom_sup_3_a

jde_bom_sup_3 %>%
  dplyr::filter(ab == "b") -> jde_bom_sup_3_b

jde_bom_sup_3_a %>% 
  dplyr::select(Component) %>% 
  dplyr::distinct()  %>% 
  dplyr::pull(Component) -> second_try_vf


exception_report_supplier %>%
  filter(item %in% second_try_vf) %>%
  pull(comp_ref) -> comp_refs_2

supplier_destination_result_2 <- sapply(comp_refs_2, find_destination_supplier, exception_report_supplier = exception_report_supplier)

supplier_destination_df_2 <- data.frame(
  comp_ref = names(supplier_destination_result_2),
  supplier = unname(supplier_destination_result_2)
)

# Clean column names
supplier_destination_df_2 %>% 
  janitor::clean_names() -> supplier_destination_df_2




jde_bom_sup_3_a %>% 
  dplyr::select(-supplier, -supplier_name) %>% 
  dplyr::left_join(supplier_destination_df_2, by = c("mpf_ref" = "comp_ref")) %>% 
  dplyr::left_join(supplier_name, by = c("supplier" = "Supplier")) %>%
  dplyr::relocate(c(supplier, supplier_name, ab, mpf_ref), .after = mpf_or_line) -> jde_bom_sup_3_a



rbind(jde_bom_sup_3_a, jde_bom_sup_3_b) -> jde_bom


jde_bom %>% 
  dplyr::select(-ab, -mpf_ref, -mpf_or_line) -> jde_bom





######################################################################################################################
##################################################### final touch ####################################################
######################################################################################################################

jde_bom %>% 
  dplyr::mutate(ref = gsub("_", "-", ref),
                comp_ref = gsub("_", "-", comp_ref)) %>% 
  dplyr::relocate(ref, comp_ref, supplier, supplier_name, Sku_Status, Category, Platform, Label, where_used_count_per_loc, where_used_count_all_loc, Business_Unit, Level, Parent_Item_Number,
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


# Supplier and Supplier Name 
jde_bom %>% 
  dplyr::mutate(supplier = ifelse(is.na(supplier), "NA", supplier)) %>% 
  dplyr::mutate(supplier_name = ifelse(is.na(supplier_name), "NA", supplier_name)) -> jde_bom








# Supplier DNRR to DNRR, NA to NA
exception_report -> exception_report_supplier


exception_report_supplier %>% 
  janitor::clean_names() %>% 
  dplyr::select(loc, item) %>% 
  dplyr::distinct() %>% 
  dplyr::mutate(supplier_ref = paste0(loc, "_", item),
                dnrr = "N") %>% 
  dplyr::select(supplier_ref, dnrr) -> exception_report_supplier

jde_bom %>% 
  dplyr::mutate(supplier_ref = comp_ref,
                supplier_ref = gsub("-", "_", supplier_ref)) %>% 
  dplyr::left_join(exception_report_supplier, by = "supplier_ref") %>%
  dplyr::mutate(dnrr = ifelse(is.na(dnrr), "Y", dnrr)) %>%
  dplyr::mutate(supplier = ifelse(supplier == "NA" & dnrr == "Y", "DNRR", 
                                  ifelse(supplier == "NA" & dnrr == "N", "NA", supplier))) %>% 
  dplyr::mutate(supplier_name = ifelse(supplier == "DNRR", "DNRR", 
                                       ifelse(supplier == "NA", "NA", supplier_name))) %>% 
  dplyr::select(-supplier_ref, -dnrr) -> jde_bom



## bring class ref
class_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Class reference (JDE).xlsx")

class_ref %>% 
  dplyr::slice(-1) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(class_number = code, class_description = description) %>% 
  dplyr::distinct(class_number, .keep_all = TRUE) -> class_ref_lookup



class_ref_lookup %>%
  mutate(
    item_type = case_when(
      as.numeric(class_number) < 500 ~ "non-commodity",
      class_number == "570" ~ "Label",
      as.numeric(class_number) >= 500 & as.numeric(class_number) < 900 & class_number != "570" ~ "packaging",
      as.numeric(class_number) > 900 ~ "commodity oil",
      class_number %in% c("BCH", "BLD", "FGT", "RPS", "SFM", "SSA", "WIP", "S09", "S08", "S04", "S01", "S02", "S03", "S06", "S05", "S10", "S07") ~ "WIP",
      class_number == "OHD" ~ "overhead",
      TRUE ~ NA_character_
    )
  ) %>% 
  mutate(item_type = ifelse(is.na(item_type), "NA", item_type)) -> class_ref_lookup_table



jde_bom %>% 
  dplyr::left_join(class_ref_lookup_table %>% 
                     rename(Commodity_Class = class_number), by = "Commodity_Class") -> jde_bom



jde_bom %>%
  dplyr::mutate(supplier = ifelse(item_type == "commodity oil" | item_type == "WIP" | item_type == "overhead", "NA", supplier)) %>%
  dplyr::mutate(supplier_name = ifelse(supplier == "DNRR", "DNRR", 
                                       ifelse(supplier == "NA", "NA", supplier_name))) %>% 
  dplyr::select(-class_description, -item_type) -> jde_bom


jde_bom %>% 
  dplyr::mutate(supplier = ifelse(is.na(supplier), "NA", supplier)) %>%
  dplyr::mutate(supplier_name = ifelse(is.na(supplier_name), "NA", supplier_name)) -> jde_bom



#####################





colnames(jde_bom)[1]<-"ref"
colnames(jde_bom)[2]<-"comp ref"
colnames(jde_bom)[3]<-"Supplier"
colnames(jde_bom)[4]<-"Supplier Name"
colnames(jde_bom)[5]<-"SKU Status"
colnames(jde_bom)[6]<-"Category"
colnames(jde_bom)[7]<-"Platform"
colnames(jde_bom)[8]<-"Label"
colnames(jde_bom)[9]<-"where used count (per loc)"
colnames(jde_bom)[10]<-"where used count (all loc)"
colnames(jde_bom)[11]<-"Business Unit"
colnames(jde_bom)[12]<-"Level"
colnames(jde_bom)[13]<-"Parent Item Number"
colnames(jde_bom)[14]<-"Parent Description"
colnames(jde_bom)[15]<-"UOM"
colnames(jde_bom)[16]<-"Net_wt"
colnames(jde_bom)[17]<-"FG On Hand"
colnames(jde_bom)[18]<-"FG Weeks on Hand"
colnames(jde_bom)[19]<-"Component"
colnames(jde_bom)[20]<-"Component Description"
colnames(jde_bom)[21]<-"Commodity Class"
colnames(jde_bom)[22]<-"UM"
colnames(jde_bom)[23]<-"Lead Time"
colnames(jde_bom)[24]<-"RM On Hand"
colnames(jde_bom)[25]<-"RM Total Weeks on Hand"
colnames(jde_bom)[26]<-"Stocking Type"
colnames(jde_bom)[27]<-"Percent Scrap"
colnames(jde_bom)[28]<-"Quantity Per"
colnames(jde_bom)[29]<-"Quantity w/ Scrap"
colnames(jde_bom)[30]<-"Unit Cost"
colnames(jde_bom)[31]<-"next 28 days open order"
colnames(jde_bom)[32]<-"mon_a fcst"
colnames(jde_bom)[33]<-"mon_b fcst"
colnames(jde_bom)[34]<-"mon_c fcst"
colnames(jde_bom)[35]<-"mon_d fcst"
colnames(jde_bom)[36]<-"mon_e fcst"
colnames(jde_bom)[37]<-"mon_f fcst"
colnames(jde_bom)[38]<-"mon_g fcst"
colnames(jde_bom)[39]<-"mon_h fcst"
colnames(jde_bom)[40]<-"mon_i fcst"
colnames(jde_bom)[41]<-"mon_j fcst"
colnames(jde_bom)[42]<-"mon_k fcst"
colnames(jde_bom)[43]<-"mon_l fcst"
colnames(jde_bom)[44]<-"mon_a dep demand"
colnames(jde_bom)[45]<-"mon_b dep demand"
colnames(jde_bom)[46]<-"mon_c dep demand"
colnames(jde_bom)[47]<-"mon_d dep demand"
colnames(jde_bom)[48]<-"mon_e dep demand"
colnames(jde_bom)[49]<-"mon_f dep demand"
colnames(jde_bom)[50]<-"mon_g dep demand"
colnames(jde_bom)[51]<-"mon_h dep demand"
colnames(jde_bom)[52]<-"mon_i dep demand"
colnames(jde_bom)[53]<-"mon_j dep demand"
colnames(jde_bom)[54]<-"mon_k dep demand"
colnames(jde_bom)[55]<-"mon_l dep demand"

writexl::write_xlsx(jde_bom, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/2025/01.14.2025/Bill of Material_01142025.xlsx")


file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/2025/01.07.2025/JDE BoM 01.07.2025.xlsx", 
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/2025/01.14.2025/JDE BoM 01.14.2025.xlsx")


# Don't forget to check Net lbs 


# After you are done with JDE
file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/2025/01.14.2025/JDE BoM 01.14.2025.xlsx", 
          "S:/Supply Chain Projects/Data Source (SCE)/JDE BoM/2025/JDE BoM 01.14.2025.xlsx", overwrite = TRUE)




