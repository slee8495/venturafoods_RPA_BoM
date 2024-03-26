jde_bom %>% filter(Business_Unit == "622")

# line 835

# exception_report -> campus column add -> do the lookup for the first value. 
# in exception report, if MPF says location numbers instead of PKG or LBL, then supplier lookup from MPF. 

exception_report %>% filter(Item == "261500100")


# =if(if(mpf_line is not number, do the vlookup from mpf), # do the vlookup from campus in exceprion report first, )



exception_report %>% 907203
  filter(Loc %in% c("622", "628", "631")) %>% filter(Item == "261500100") -> a

# a %>% mutate(MPF_or_Line = as.double(MPF_or_Line) %>% mutate(ifelse(!is.na(MPF_or_Line)) # create mpf_ref))
                                                             
# do the vlookup for jde_bom based on mpf_ref. NA remains NA. 


## We want to try this for Supplier and Lead time columns. 

