# This script queries ofm postcensal and intercensal Total Housing Unit data for the Puget Sound Region and exports to .xlsx

library(openxlsx)
library(data.table)
library(plotly) # version 4.0 and up
library(magrittr)
library(stringr)
library(RColorBrewer)

## user inputs--------------------------------
# time period
all <- 1
present <- 0
interc.0010 <- 0
interc.9000 <- 0

# format
excel <- 0
plotly <- 1
plot.ann.est <- 1
plot.ann.chg <- 1

# directory and files
indir <- "J:/Staff/Christy/OFM/2016/"
outdir <- "C:/Users/Christy/Desktop/ofm/results"
my.filename <- "ofm_housing_2010_present_cps.xlsx"

# read data files 
ofmApr1.present.file <- file.path(indir, "ofm_april1_housing.xlsx")
ofmApr1.00.10.file <- file.path(indir, "ofm_april1_intercensal_estimates_2000-2010.xlsx")
ofmApr1.90.00.file <- file.path(indir, "ofm_april1_intercensal_estimates_1990-2000.xlsx")

present.yr <- 2016
present.years <- seq(2010, present.yr)
inter.years2 <- seq(2000, 2010)
inter.years1 <- seq(1990, 2000)
county <- c("King", "Kitsap", "Pierce", "Snohomish")

## functions--------------------------------
# function to make sequence of years grep ready
years.as.regex <- function(years){
  x <- toString(years)
  y <- gsub(',', '|', x)
  z <- gsub(" ", "", y, fixed = TRUE)
}

# function to query ofm file
query.table <- function(table, year){
  select.cols <- (colnames(table)[grepl(paste0("Filter|County|Jurisdiction|", year),names(table))])
  select.cols2 <- grep("Filter|County|Jurisdiction|Total", select.cols, value = TRUE)
  table <- table[,colnames(table) %in% select.cols2]
  table <- table[table$County %in% county, ]
  table <- table[(table$Filter %in% c(1)), ]
}

# function to prep ofm present table
ofm_april1_present <- function (file, years){
  table <- read.xlsx(file, sheet = "Housing Units", startRow = 4, colNames = TRUE, rowNames = FALSE, rows = c(4:455), cols = NULL)
  year <- years.as.regex(years)
  table <- query.table(table, year)
  colnames(table)[4:ncol(table)] <- lapply(years, function(x) paste0("yr", x))
  table[4:ncol(table)] <- lapply(table[4:ncol(table)], function(x) as.integer(as.character(x)))
  return(table)
}

# function to prep ofm intercensal table
ofm_april1_intercensal <- function (file, years){
  table <- read.xlsx(file, sheet = "Total Housing", startRow = 1, colNames = TRUE, rowNames = FALSE, rows = c(1:455), cols = NULL)
  year <- years.as.regex(years)
  table <- query.table(table, year)
  colnames(table)[4:ncol(table)] <- lapply(years, function(x) paste0("yr", x))
  table[4:ncol(table)] <- lapply(table[4:ncol(table)], function(x) as.integer(as.character(x)))
  return(table)
}

# function to calculate annual change
ann.change <- function(table){
  dt <- as.data.table(table)
  dt.colnames <- colnames(dt)
  dt.colnames1 <- tail(dt.colnames, -3)
  gx <- tail(dt.colnames1, -1)
  gy <- head(dt.colnames1, -1)
  gn <- paste0("D[",gy,":",gx,"]")
  dt2 <- dt[,(gn) := mapply(function(x,y).SD[[x]]-.SD[[y]], gx, gy, SIMPLIFY=FALSE)]
  dt.colnames2 <- colnames(dt2)
  select.dt.colnames2 <- grep("F|C|J|D", dt.colnames2, value = TRUE)
  dt3 <- dt2[,select.dt.colnames2, with=FALSE]
}

# function to calculate regional total
region.total <- function(table){
  dt <- as.data.table(table)
  dt.colnames <- colnames(dt)
  select.dt.colnames <- grep("yr|D", dt.colnames, value = TRUE)
  rdt <- dt[, lapply(.SD, sum, na.rm=TRUE), .SDcols=c(select.dt.colnames)]
  rdt[,':='(Filter = 5, County.Name = 'Central Puget Sound Region', Jurisdiction = 'Central Puget Sound Region')]
  setcolorder(rdt, c("Filter", "County.Name", "Jurisdiction", select.dt.colnames))
}

# function to append worksheet
append.worksheet <- function(table, workbook, sheetName){
  addWorksheet(workbook, sheetName)
  writeData(workbook, sheet = sheetName, x = table, colNames = TRUE)
  saveWorkbook(workbook, paste0(outdir, my.filename), overwrite = TRUE )
}

# Call functions--------------------------------
if (all == 1){
  t1 <- ofm_april1_intercensal(ofmApr1.90.00.file, inter.years1)
  t2 <- ofm_april1_intercensal(ofmApr1.00.10.file, inter.years2)
  t3 <- ofm_april1_present(ofmApr1.present.file, present.years)
  
  t1 <- t1[,c(1:ncol(t1)-1)]
  table <- merge(t1, t2, by = c("Filter", "County.Name", "Jurisdiction"))
  table <- table[,c(1:ncol(table)-1)]
  table <- merge(table, t3, by.x = c("Filter", "County.Name", "Jurisdiction"), by.y = c("Filter", "County", "Jurisdiction"))
  if (excel == 1){
    # write table to workbook
    wb <- createWorkbook(paste0(outdir, my.filename))
    addWorksheet(wb, "annual estimates")
    writeData(wb, sheet = "annual estimates", x = table, colNames = TRUE)
    print("exported annual estimates")
    
    #write reg.table to workbook
    reg.table <- region.total(table)
    append.worksheet(reg.table, wb, "reg annual estimates")
    print("exported regional annual estimates")
    
    # write annual change to workbook
    ann.change.table <- ann.change(table)
    append.worksheet(ann.change.table, wb, "annual change")
    print("exported annual change estimates")
    
    #write reg.ann.table to workbook
    reg.ann.change.table <- region.total(ann.change.table)
    append.worksheet(reg.ann.change.table, wb, "reg annual change")
    print("exported regional annual change estimates")
  }
} else if (present == 1){
  table <- ofm_april1_present(ofmApr1.present.file, present.years)
} else if (interc.0010 == 1){
  table <- ofm_april1_intercensal(ofmApr1.00.10.file, inter.years2)
} else if (interc.9000 == 1){
  table <- ofm_april1_intercensal(ofmApr1.90.00.file, inter.years1)
}

## Plots--------------------------------------------
# plot annual county estimates as stacked bar chart
if (plotly == 1 & plot.ann.est == 1){ 
  # transform table
  select.yrcols.ind <-grep(paste0("yr"),names(table))
  ptable <- NULL
  for (i in 1:length(select.yrcols.ind)){
    t <- NULL
    t <- table[,c(1:3, select.yrcols.ind[i])] 
    t$year <- str_match(tail(colnames(t), 1), "(\\d+)")[,1]
    colnames(t)[ncol(t)-1] <- "estimate"
    ifelse(is.null(ptable), ptable <- t, ptable <- rbind(ptable, t))
  }
  
  # plot table
  p <- plot_ly(ptable,
               x = ~year,
               y = ~estimate,
               color = ~County.Name,
               colors = c("#377eb8", "#ff7f00", "#4daf4a", "#e41a1c"),
               type = 'bar'
  )%>%
    layout(title = "OFM Total Housing Unit Estimate",
           font = list(family="Segoe UI", size = 12.5),
           barmode = 'stack')
  
  print(p) 
}

# plot annual change by county (subplot)
if (plotly == 1 & plot.ann.chg == 1){ 
  # transform table
  ann.change.table <- ann.change(table) %>% as.data.frame()
  select.yrcols.ind <-grep(paste0("D"),names(ann.change.table))
  ptable <- NULL
  for (i in 1:length(select.yrcols.ind)){
    t <- NULL
    t <- ann.change.table[,c(1:3, select.yrcols.ind[i])] 
    t$year <- str_match(tail(colnames(t), 1), "(\\d+):yr(\\d+)")[,1]
    t$year <- gsub(":yr", "-", t$year)
    colnames(t)[ncol(t)-1] <- "estimate"
    ifelse(is.null(ptable), ptable <- t, ptable <- rbind(ptable, t))
  }
  
  ptable$id <- as.integer(factor(ptable$County.Name))
  # plot table
  one_plot <- function(dat){
    plot_ly(dat,
            x = ~year,
            y = ~estimate,
            split = ~County.Name,
            type = 'bar')
  }
  
  # plot
  p <- ptable %>%
    group_by(County.Name) %>%
    do(p = one_plot(.)) %>%
    subplot(nrows = 2) %>%
    layout(xaxis = list(domain = c(0, .45), showgrid=TRUE),
           yaxis = list(domain = c(.55, 1)),
           
           xaxis2 = list(domain = c(.55, 1), showgrid=TRUE),
           yaxis2 = list(domain = c(.55, 1)),
           
           xaxis3 = list(domain = c(0, .45), showgrid=TRUE),
           yaxis3 = list(domain = c(0, .45)),
           
           xaxis4 = list(domain = c(.55, 1), showgrid=TRUE),
           yaxis4 = list(domain = c(0, .45)),
           
           title = "OFM Housing Unit Growth",
           font = list(family="Segoe UI", size = 11),
           margin = list(l=100, b=50, t=50, r=100))
  
  print(p) 
}

