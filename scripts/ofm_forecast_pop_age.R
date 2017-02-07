# This script compiles OFM 2012 population by age projections and computes an estimate of the 
# 15-17 age cohort based on the share of persons age 15-17 from the 2014 Census national projection
# for years 2025 and 2040.

library(openxlsx)
library(dplyr)
library(data.table)

# function to assemble OFM tables for each geography
assemble.juris.table <- function(jurisdiction, start.row){
  master.file <- file.path(indir2, ofm.file)
  section1 <- read.xlsx(master.file, sheet = 1, startRow = as.integer(start.row), colNames = TRUE, rowNames = FALSE, rows = c(as.integer(start.row):as.integer(start.row + 20)), cols = NULL)
  section2 <- read.xlsx(master.file, sheet = 1, startRow = as.integer(start.row + 22), colNames = TRUE, rowNames = FALSE, rows = c(as.integer(start.row + 22): as.integer(start.row + 42)), cols = NULL)
  section3 <- read.xlsx(master.file, sheet = 1, startRow = as.integer(start.row + 44), colNames = TRUE, rowNames = FALSE, rows = c(as.integer(start.row + 44): as.integer(start.row + 64)), cols = c(1:2))
  table <- bind_cols(section1, section2, section3)
  colnames(table)[1] <- "agegroups"
  select.cols <- (colnames(table)[grepl(paste0("groups|\\d+{4}"), names(table))])
  table <- table[,colnames(table) %in% select.cols]
  colnames(table)[2:ncol(table)] <- paste0("tot", select.cols[2:length(select.cols)])
  table <- table[-c(1), ]
  table$jurisdiction <- jurisdiction
  table %>% mutate_each(funs(as.numeric), starts_with("tot"))
} 

# function to append worksheet
append.worksheet <- function(table, workbook, sheetName){
  addWorksheet(workbook, sheetName)
  writeData(workbook, sheet = sheetName, x = table, colNames = TRUE)
  saveWorkbook(workbook, file.path(indir, my.filename), overwrite = TRUE )
}

indir <- "J:/Staff/Christy/OFM/forecasts/requests/jean"
indir2 <- "J:/Staff/Christy/OFM/forecasts"
my.filename <- "youth_senior_pop_projections.xlsx"

census.file <- file.path(indir, "NP2014_D1.csv")
ofm.file <- "gma2012_cntyage_med.xlsx"

forecast.years <- c(2025, 2040)
censusdt<- as.data.table(read.table(census.file, header = TRUE, sep = ","))
cols <- c("year", "total_pop", "pop_15", "pop_16", "pop_17", "pop_18", "pop_19")
censusdt.select <- censusdt[origin == 0 & race == 0 & sex == 0 & year %in% forecast.years, cols, with = FALSE 
                            ][, pop_15_19 := rowSums(.SD), .SDcols = 3:7
                              ][, pop_15_17 := rowSums(.SD), .SDcols = 3:5
                                ][, share_pop15_17 := pop_15_17/pop_15_19
                                  ][, year := paste0("tot", year)]

king <- assemble.juris.table("King County", 1176)
kitsap <- assemble.juris.table("Kitsap County", 1245)
pierce <- assemble.juris.table("Pierce County", 1866)
snohomish <- assemble.juris.table("Snohomish County", 2142)
region <- bind_rows(king, kitsap, pierce, snohomish)

region.youth <- as.data.table(region)
ry.m <- melt(region.youth, id.vars = c("agegroups", "jurisdiction"), measure.vars = c("tot2010", "tot2015", "tot2020",  "tot2025", "tot2030", "tot2035","tot2040"))
ry.m.select <- ry.m[agegroups == "15-19" & variable %in% paste0("tot", forecast.years),]
setkey(ry.m.select, variable)[censusdt.select, pop15_17 := round(value*share_pop15_17, 0)]
youth <- ry.m.select[,.(agegroups = "15-17", jurisdiction, variable, value = pop15_17),]

dts <- list(ry.m, youth)
region.pop.dt <- rbindlist(dts, use.names = TRUE, fill = TRUE)
format.region.pop.dt <- dcast.data.table(region.pop.dt, agegroups + jurisdiction ~ variable)
format.region.pop.dt2 <- format.region.pop.dt[agegroups %in% c('5-9', '10-14', '15-17', '65-69', '70-74', '75-79', '80-84', '85+'), .(agegroups, jurisdiction, tot2025, tot2040),]

youth.age <- c('5-9', '10-14', '15-17')
plus.65 <- c('65-69', '70-74', '75-79', '80-84', '85+')

format.region.pop.dt2 <- format.region.pop.dt2[agegroups %in% youth.age, category := "youth",
                                               ][agegroups %in% plus.65, category := "65+",]
format.region.pop.dt3 <- format.region.pop.dt2[,lapply(.SD, sum), by = list(jurisdiction, category), .SDcols = paste0("tot", forecast.years)]
format.region.pop.dt4 <- format.region.pop.dt2[agegroups == '85+',][,category := '85+']
dt.agg <- rbindlist(list(format.region.pop.dt3, format.region.pop.dt4), use.names = TRUE, fill = TRUE)
dt.county <- dt.agg[,agegroups := NULL]
dt.region <- dt.county[, lapply(.SD, sum), by = category, .SDcols = paste0("tot", forecast.years)
                       ][,jurisdiction := 'Puget Sound Region']
dt <- rbindlist(list(dt.county, dt.region), use.names = TRUE, fill = TRUE)

# write dt to workbook
wb <- createWorkbook(file.path(indir, my.filename))
addWorksheet(wb, "Youth and Seniors")
writeData(wb, sheet = "Youth and Seniors", x = dt, colNames = TRUE)
print("exported Youth and Seniors")

# append region to workbook
append.worksheet(region, wb, "OFM proj for CPS")
print("exported regional OFM projections")

# append censusdt.select to workbook
append.worksheet(censusdt.select, wb, "Census natl proj for 15_17")
print("exported Census national projections for youth")