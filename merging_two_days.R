glimpse(HP_010318_cleaned)

HP_020318_cleaned$bill_number <- as.integer(HP_020318_cleaned$bill_number)
HP_020318_cleaned$laundry_fee <- as.integer(HP_020318_cleaned$laundry_fee)

#glimpse(HP_020318_cleaned)

two_days <- HP_010318_cleaned %>% 
  bind_rows(HP_020318_cleaned)


