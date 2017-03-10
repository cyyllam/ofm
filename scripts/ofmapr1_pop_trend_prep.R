# This script prepares a given year's ofm Apr 1 pop estimates and annual annexation pop estimates up to the given Apr 1 date.
library(openxlsx)
library(magrittr)

## User inputs------------------------------
year <- "2016"
indir <- "J:/Staff/Christy/OFM"
indir.path <- file.path(indir, year)
my.filename <- paste0("poptrend_dataprep", year, ".xlsx")
outdir <- "C:/Users/Christy/Desktop/ofm/results"
county <- c("King", "Kitsap", "Pierce", "Snohomish")
newcolname <- paste0("pop", year)

## functions------------------------------
# function to query ofm pop file
query.table <- function(table, attributes, filternum){
  select.cols <- (colnames(table)[grepl(paste0("Filter|County|Jurisdiction|", attributes), names(table))])
  table <- table[,colnames(table) %in% select.cols]
  table <- table[table$County %in% county, ]
  table <- table[(table$Filter %in% filternum), ]
}

# function to append worksheet
append.worksheet <- function(table, workbook, sheetName){
  addWorksheet(workbook, sheetName)
  writeData(workbook, sheet = sheetName, x = table, colNames = TRUE)
  saveWorkbook(workbook, file.path(outdir, my.filename), overwrite = TRUE )
}

# function to change estimates colname and convert to int
change.name.type <- function(table, newcolname){
  colnames(table)[ncol(table)] <- newcolname
  table[ncol(table)] <- lapply(table[ncol(table)], function(x) as.integer(as.character(x)))
  return(table)
}

# function to aggregate multicnty cities' records
agg.multicnty.cities <- function(table, attribute){
  t <- table[grepl("[(]part[)]", table[, "Jurisdiction"]), ] 
  t2 <- aggregate(t[ , ncol(t)], by = list(t$Jurisdiction), FUN = sum)
  colnames(t2) <- c("Jurisdiction", attribute)
  t2$Jurisdiction <- gsub("[(]part[)]", "(all)", t2$Jurisdiction)
  t2$Filter <- "4"
  t2$County <- "zMulti"
  return(t2)
}

## create and export tables------------------------------
# ofm april 1 population final, county 
ofmApr1.pop.file <- file.path(indir.path, "ofm_april1_population_final.xlsx")
pop.table <- read.xlsx(ofmApr1.pop.file, sheet = "Population", startRow = 5, colNames = TRUE, rowNames = FALSE, cols = NULL)
cnty.table <- query.table(pop.table, year, seq(1:3)) %>% change.name.type(newcolname)

# ofm april 1 population final, cities & cities in multi-counties
cities.table <-query.table(pop.table, year, 4) %>% change.name.type(newcolname)
multicnty.cities <- agg.multicnty.cities(cities.table, newcolname)
cities.table <- rbind(cities.table, multicnty.cities)

# ofm annex summary
ofm.annexsum.file <- file.path(indir.path, "annex_summary.xlsx")
annex.table <- read.xlsx(ofm.annexsum.file, sheet = year, startRow = 1, colNames = TRUE, rowNames = FALSE, cols = NULL)
annexation.table <- query.table(annex.table, "Annexation[.]Population", c(1,4))
annexation.table$Annexation.Population <- gsub("\\.", 0, annexation.table$Annexation.Population)
annexation.table[ncol(annexation.table)] <- lapply(annexation.table[ncol(annexation.table)], function(x) as.integer(as.character(x)))

multicnty.annex <- agg.multicnty.cities(annexation.table, "Annexation.Population")
annexation.table <- rbind(annexation.table, multicnty.annex)

# write table to workbook
wb <- createWorkbook(file.path(outdir, my.filename))
addWorksheet(wb, "counties")
writeData(wb, sheet = "counties", x = cnty.table, colNames = TRUE)
print("counties")

#write cities.table to workbook
append.worksheet(cities.table, wb, "cities")
print("exported cities")

#write annexation.table to workbook
append.worksheet(annexation.table, wb, "annexations")
print("exported annexations")
