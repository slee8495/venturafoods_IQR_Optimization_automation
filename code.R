library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(officer)
library(openxlsx)
library(lubridate)
library(magrittr)
library(skimr)
library(bizdays)
library(janitor) 

##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################



# Read in the data
iqr_fg <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.16.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 07.16.2024.xlsx",
                                  sheet = "Location FG")


iqr_fg[-1:-2, ] -> iqr_fg
colnames(iqr_fg) <- iqr_fg[1, ]
iqr_fg[-1, ] -> iqr_fg



inventory_model_data_fg <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx",
                                   col_names = FALSE, sheet = "Fin Goods")


inventory_model_data_fg[-1:-7, ] -> inventory_model_data_fg
colnames(inventory_model_data_fg) <- inventory_model_data_fg[1, ]
inventory_model_data_fg[-1, ] -> inventory_model_data_fg

iqr_rm <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Weekly Report run/2024/08.27.2024/Raw Material Inventory Health (IQR) NEW TEMPLATE - 08.27.2024.xlsx",
                                  sheet = "RM data")

iqr_rm[-1:-2, ] -> iqr_rm
colnames(iqr_rm) <- iqr_rm[1, ]
iqr_rm[-1, ] -> iqr_rm 


inventory_model_data_rm <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Raw Material LIVE.xlsx",
                              sheet = "Sheet1")

inventory_model_data_rm[-1:-5,] -> inventory_model_data_rm
colnames(inventory_model_data_rm) <- inventory_model_data_rm[1, ]
inventory_model_data_rm[-1, ] -> inventory_model_data_rm


#################################################### Finished Goods ############################################################################################    

### 1. Has on hand inventory (useable, soft hold, hard hold all included)
iqr_fg %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(usable = as.double(usable),
                quality_hold = as.double(quality_hold),
                soft_hold = as.double(soft_hold)) %>% 
  dplyr::mutate(inventory = usable + quality_hold + soft_hold) %>% 
  dplyr::filter(inventory > 0) %>% 
  dplyr::select(loc, campus, item_2, base) %>% 
  dplyr::distinct_all() -> has_on_hand_inventory_fg


### 2. Zero on hand, but has open customer orders & branch transfer for next 30 days
iqr_fg %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(usable = as.double(usable),
                quality_hold = as.double(quality_hold),
                soft_hold = as.double(soft_hold)) %>% 
  dplyr::mutate(inventory = usable + quality_hold + soft_hold) %>%
  dplyr::filter(inventory == 0) %>%
  dplyr::filter(cust_ord_in_next_28_days > 0 | mfg_cust_ord_in_next_28_days > 0) %>% 
  dplyr::select(loc, campus, item_2, base) %>%
  dplyr::distinct_all() -> zero_on_hand_has_open_orders_fg


### 3. Zero on hand, zero open customer orders, but has forecast for next 12 months 
iqr_fg %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(usable = as.double(usable),
                quality_hold = as.double(quality_hold),
                soft_hold = as.double(soft_hold)) %>% 
  dplyr::mutate(inventory = usable + quality_hold + soft_hold) %>%
  dplyr::filter(inventory == 0) %>%
  dplyr::filter(cust_ord_in_next_28_days == 0 | mfg_cust_ord_in_next_28_days == 0) %>% 
  dplyr::filter(total_forecast_next_12_months > 0 | total_mfg_forecast_next_12_months > 0) %>% 
  dplyr::select(loc, campus, item_2, base) %>%
  dplyr::distinct_all() -> zero_on_hand_zero_open_orders_has_forecast_fg


### 4. None of the above but show ACTIVE in JDE
iqr_fg %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(ref) %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) -> iqr_fg_4 

inventory_model_data_fg %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(ship_ref, from_jde) %>% 
  dplyr::rename(ref = ship_ref) %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) -> inventory_model_4_fg

  
dplyr::left_join(iqr_fg_4, inventory_model_4_fg, by = "ref") -> criteria_4_a_fg
dplyr::anti_join(inventory_model_4_fg %>% dplyr::filter(from_jde == "ACTIVE"), 
                                      iqr_fg_4, by = "ref") -> criteria_4_b_fg

dplyr::bind_rows(criteria_4_a_fg, criteria_4_b_fg) %>% 
  dplyr::filter(from_jde == "ACTIVE") %>% 
  dplyr::select(ref) -> active_items_fg



## Conclusion
dplyr::bind_rows(has_on_hand_inventory_fg, 
                 zero_on_hand_has_open_orders_fg, 
                 zero_on_hand_zero_open_orders_has_forecast_fg, 
                 active_items_fg) %>%
  dplyr::distinct(ref) -> final_data_fg





#################################################### Raw Materials ############################################################################################


# 1. Has inventory (useable, soft hold, hard hold all included)
iqr_rm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(usable = as.double(usable),
                quality_hold = as.double(quality_hold),
                soft_hold = as.double(soft_hold)) %>% 
  dplyr::mutate(inventory = usable + quality_hold + soft_hold) %>% 
  dplyr::filter(inventory > 0) %>% 
  tidyr::separate(loc_sku, into = c("loc", "item_2"), sep = "-") %>% 
  dplyr::select(loc, item_2) %>% 
  dplyr::distinct_all() -> has_on_hand_inventory_rm


# 2. Zero inventory but has dependent demand for next 6 months
iqr_rm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(usable = as.double(usable),
                quality_hold = as.double(quality_hold),
                soft_hold = as.double(soft_hold)) %>%
  dplyr::mutate(inventory = usable + quality_hold + soft_hold) %>%
  dplyr::filter(inventory == 0) %>%
  dplyr::filter(total_dep_demand_next_6_months > 0) %>%
  tidyr::separate(loc_sku, into = c("loc", "item_2"), sep = "-") %>% 
  dplyr::select(loc, item_2) %>% 
  dplyr::distinct_all() -> zero_on_hand_has_dependent_demand_rm

#  3. Zero inventory, no dependent demand for next 6 months but show ACTIVE in JDE
iqr_rm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::rename(ref = loc_sku) %>% 
  dplyr::select(ref) %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) -> iqr_rm_3 

inventory_model_data_rm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(loc_sku, product_status) %>% 
  dplyr::rename(ref = loc_sku) %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) -> inventory_model_3_rm


dplyr::left_join(iqr_rm_3, inventory_model_3_rm, by = "ref") -> criteria_4_a_rm
dplyr::anti_join(inventory_model_3_rm %>% dplyr::filter(product_status == "ACTIVE"), 
                 iqr_rm_3, by = "ref") -> criteria_4_b_rm

dplyr::bind_rows(criteria_4_a_rm, criteria_4_b_rm) %>% 
  dplyr::filter(product_status == "ACTIVE") %>% 
  dplyr::select(ref) -> active_items_rm


## Conclusion
dplyr::bind_rows(has_on_hand_inventory_rm, 
                 zero_on_hand_has_dependent_demand_rm, 
                 active_items_rm) %>%
  dplyr::distinct(ref) -> final_data_rm




