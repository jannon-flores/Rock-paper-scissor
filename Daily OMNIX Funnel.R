rm(list=ls(all=TRUE))

library(xtable)
library(data.table)
library(httr) ## addresses firewall issues when connecting to GA 
library(openxlsx)
library(dplyr)
library(RDCOMClient)
library(formattable)
library(RPostgreSQL)
library(lubridate)

library(rsconnect)
library(anytime)

setwd('C:/Users/AnalyticsReporting/Desktop')
#setwd('C:/Users/AnalyticsReporting/Desktop')

library(googleAuthR)
my_client_id <- '632523129845-vqo8sghg047gh1hcr1j0rg7q51mlttt3.apps.googleusercontent.com'
my_client_secret <- 'NF30t1DQkQ8pyjKYUvJ8mIR9'

options(googleAuthR.client_id = my_client_id)
options(googleAuthR.client_secret = my_client_secret)

library(googleAnalyticsR)
gar_auth_service("C:/Users/AnalyticsReporting/Documents/auth/cebu-pacific-account.json")



## define the segments of the goal funnel steps
seg0 <- segment_element("ga:pagePath",
                        operator = "EXACT",
                        type = "DIMENSION",
                        caseSensitive = NULL,
                        expressions = "/",
                        scope = "SESSION")

seg1 <- segment_element("ga:pagePath",
                        operator = "REGEXP",
                        type = "DIMENSION",
                        caseSensitive = NULL,
                        expressions = "/select-flight",
                        scope = "SESSION"
                        #,matchType = "IMMEDIATELY_PRECEDES"
)

seg2 <- segment_element("ga:pagePath",
                        operator = "REGEXP",
                        type = "DIMENSION",
                        caseSensitive = NULL,
                        expressions = "/guest-details",
                        scope = "SESSION"
                        #,matchType = "IMMEDIATELY_PRECEDES"
)

seg3 <- segment_element("ga:pagePath",
                        operator = "REGEXP",
                        type = "DIMENSION",
                        caseSensitive = NULL,
                        expressions = "/add-ons",
                        scope = "SESSION"
                        #,matchType = "IMMEDIATELY_PRECEDES"
)

seg4 <- segment_element("ga:pagePath", 
                        operator = "REGEXP", 
                        type = "DIMENSION", 
                        caseSensitive = NULL,
                        expressions = "/payment", 
                        scope = "SESSION"
                        #,matchType = "IMMEDIATELY_PRECEDES"
)

# seg5 <- segment_element("ga:pagePath",
# operator = "REGEXP",
# type = "DIMENSION",
# caseSensitive = NULL,
# expressions = "/Booking/PostCommit",
# scope = "SESSION"
# ,matchType = "IMMEDIATELY_PRECEDES"
# )

seg5 <- segment_element("ga:pagePath",
                        operator = "PARTIAL",
                        type = "DIMENSION",
                        caseSensitive = NULL,
                        expressions = "/confirmation?state_id=2001",
                        scope = "SESSION"
                        #,matchType = "IMMEDIATELY_PRECEDES"
)


## usesequence vector to ensure funneling of the goal steps
seg_svs_0 <- segment_vector_simple(list(list(seg0)))
seg_svs_1 <- segment_vector_sequence(list(list(seg1)))
seg_svs_2 <- segment_vector_sequence(list(list(seg1),list(seg2)))
seg_svs_3 <- segment_vector_sequence(list(list(seg1),list(seg2),list(seg3)))
seg_svs_4 <- segment_vector_sequence(list(list(seg1),list(seg2),list(seg3),list(seg4)))
seg_svs_5 <- segment_vector_sequence(list(list(seg1),list(seg2),list(seg3),list(seg4),list(seg5)))
#seg_svs_6 <- segment_vector_sequence(list(list(seg1),list(seg2),list(seg3),list(seg4),list(seg5),list(seg6)))

## name each goal funnel step accordingly
seg_g4_0 <- segment_ga4("00 Homepage", session_segment = segment_define(list(seg_svs_0)))
seg_g4_1 <- segment_ga4("01 New - Select", session_segment = segment_define(list(seg_svs_1)))
seg_g4_2 <- segment_ga4("02 New - Guest Details", session_segment = segment_define(list(seg_svs_2)))
seg_g4_3 <- segment_ga4("03 New - Extras", session_segment = segment_define(list(seg_svs_3)))
seg_g4_4 <- segment_ga4("04 New - Payment", session_segment = segment_define(list(seg_svs_4)))
seg_g4_5 <- segment_ga4("05 New - Purchase Ticket", session_segment = segment_define(list(seg_svs_5)))
#seg_g4_6 <- segment_ga4("06 New - Purchase Ticket", session_segment = segment_define(list(seg_svs_6)))

## no. of dates to cover with respect to the start_date
n <- 1

## define start_date as current date and end_date based on n 
end_date <- Sys.Date() - 1 # current date -1 to ensure that complete data will be fetched
end_date <- as.Date(end_date) # ensuring date format
start_date <- end_date
start_date <- as.Date(start_date)
start_date
end_date
#start_date - end_date

## set the view ids for proper fetching of data
ga_vidf <- "211961643" # for Aggregated Reports
ga_vidd <- "211961643" # for Aggregated Reports(Desktop Only)
ga_vidm <- "211961643" # for Aggregated Reports(Mobile Only)
ga_vidt <- "211961643" # for Aggregated Reports(Tablet Only)

## create a vector for the report view IDs
gavids <- c(ga_vidf,ga_vidd,ga_vidm,ga_vidt)

## query for the aggregated (ALL DEVICES) reporting
gad_01_0 <- google_analytics(viewId = gavids[1],
                             date_range = c(start = start_date, end = end_date),
                             segments = seg_g4_0, #c(seg_g4_0, seg_g4_1,seg_g4_2,seg_g4_3),
                             metrics = "ga:sessions",
                             dimensions = c("date","segment"),
                             max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                             #,anti_sample = TRUE # no sampling will be made
)

gad_01_1 <- google_analytics(viewId = gavids[1],
                             date_range = c(start = start_date, end = end_date),
                             segments = seg_g4_1, #c(seg_g4_0, seg_g4_1,seg_g4_2,seg_g4_3),
                             metrics = "ga:sessions",
                             dimensions = c("date","segment"),
                             max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                             #,anti_sample = TRUE # no sampling will be made
)


gad_01_2 <- google_analytics(viewId = gavids[1],
                             date_range = c(start = start_date, end = end_date),
                             segments = seg_g4_2, #c(seg_g4_0, seg_g4_1,seg_g4_2,seg_g4_3),
                             metrics = "ga:sessions",
                             dimensions = c("date","segment"),
                             max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                             #,anti_sample = TRUE # no sampling will be made
)

gad_01_3 <- google_analytics(viewId = gavids[1],
                             date_range = c(start = start_date, end = end_date),
                             segments = seg_g4_3, #c(seg_g4_0, seg_g4_1,seg_g4_2,seg_g4_3),
                             metrics = "ga:sessions",
                             dimensions = c("date","segment"),
                             max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                             #,anti_sample = TRUE # no sampling will be made
)

gad_01 <- rbind(gad_01_0,gad_01_1,gad_01_2,gad_01_3)

gad_02_4 <- google_analytics(viewId = gavids[1],
                             date_range = c(start = start_date, end = end_date),
                             segments = seg_g4_4,#c(seg_g4_4, seg_g4_5,seg_g4_6),
                             metrics = "ga:sessions",
                             dimensions = c("date","segment"),
                             max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                             #,anti_sample = TRUE
)

gad_02_5 <- google_analytics(viewId = gavids[1],
                             date_range = c(start = start_date, end = end_date),
                             segments = seg_g4_5,#c(seg_g4_4, seg_g4_5,seg_g4_6),
                             metrics = "ga:sessions",
                             dimensions = c("date","segment"),
                             max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                             #,anti_sample = TRUE
)

# gad_02_6 <- google_analytics(viewId = gavids[1],
# date_range = c(start = start_date, end = end_date),
# segments = seg_g4_6,#c(seg_g4_4, seg_g4_5,seg_g4_6),
# metrics = "ga:sessions",
# dimensions = c("date","segment")
# ,max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
#,anti_sample = TRUE
# )

gad_02 <- rbind(gad_02_4,gad_02_5)

## combine all segments into one df
gad_all <- rbind(gad_01, gad_02)

## initialize new column with ALL for merging later with below query
gad_all$deviceCategory <- "ALL"

## query for the per DEVICE (Desktop, Mobile and Tablet) reporting by GROUPING
gad_pd_01_0 <- google_analytics(viewId = gavids[1],
                                date_range = c(start = start_date, end = end_date),
                                segments = seg_g4_0,#c(seg_g4_0,seg_g4_1,seg_g4_2,seg_g4_3),
                                metrics = "ga:sessions",
                                dimensions = c("date","segment", "ga:deviceCategory"),
                                max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                                #,anti_sample = TRUE
)


gad_pd_01_1 <- google_analytics(viewId = gavids[1],
                                date_range = c(start = start_date, end = end_date),
                                segments = seg_g4_1,#c(seg_g4_0,seg_g4_1,seg_g4_2,seg_g4_3),
                                metrics = "ga:sessions",
                                dimensions = c("date","segment", "ga:deviceCategory"),
                                max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                                #,anti_sample = TRUE
)


gad_pd_01_2 <- google_analytics(viewId = gavids[1],
                                date_range = c(start = start_date, end = end_date),
                                segments = seg_g4_2,#c(seg_g4_0,seg_g4_1,seg_g4_2,seg_g4_3),
                                metrics = "ga:sessions",
                                dimensions = c("date","segment", "ga:deviceCategory"),
                                max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                                #,anti_sample = TRUE
)

gad_pd_01_3 <- google_analytics(viewId = gavids[1],
                                date_range = c(start = start_date, end = end_date),
                                segments = seg_g4_3,#c(seg_g4_0,seg_g4_1,seg_g4_2,seg_g4_3),
                                metrics = "ga:sessions",
                                dimensions = c("date","segment", "ga:deviceCategory"),
                                max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                                #,anti_sample = TRUE
)

gad_pd_01 <- rbind(gad_pd_01_0,gad_pd_01_1,gad_pd_01_2,gad_pd_01_3)


gad_pd_02_4 <- google_analytics(viewId = gavids[1],
                                date_range = c(start = start_date, end = end_date),
                                segments = seg_g4_4,#c(seg_g4_4, seg_g4_5,seg_g4_6),
                                metrics = "ga:sessions",
                                dimensions = c("date","segment", "ga:deviceCategory"),
                                max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                                #,anti_sample = TRUE
)

gad_pd_02_5 <- google_analytics(viewId = gavids[1],
                                date_range = c(start = start_date, end = end_date),
                                segments = seg_g4_5,#c(seg_g4_4, seg_g4_5,seg_g4_6),
                                metrics = "ga:sessions",
                                dimensions = c("date","segment", "ga:deviceCategory"),
                                max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
                                #,anti_sample = TRUE
)

# gad_pd_02_6 <- google_analytics(viewId = gavids[1],
# date_range = c(start = start_date, end = end_date),
# segments = seg_g4_6,#c(seg_g4_4, seg_g4_5,seg_g4_6),
# metrics = "ga:sessions",
# dimensions = c("date","segment", "ga:deviceCategory"),
# max = 3000 # eq. 24 for a given single day (6 segments x 4 devices)
#,anti_sample = TRUE
# )

gad_pd_02 <- rbind(gad_pd_02_4,gad_pd_02_5)

gad_pd_all <- rbind(gad_pd_01, gad_pd_02)

## merge the two data frames for ALL and per Device
gad_full <- rbind(gad_all, gad_pd_all)

## back-up queried data 
gad_all_bkp <- gad_full

head(gad_full,20)
tail(gad_full)

gad_full[gad_full$segment=="05 New - Purchase Ticket",]

## sort the final data frame by segment and then re-index
gad_full <- as.data.table(gad_full)
gad_full <- gad_full %>% arrange(date, deviceCategory, segment)
head(gad_full,20)

#gad_full <- gad_full[order(gad_full$date, gad_full$deviceCategory, gad_full$segment), na.last=NA),]
#row.names(gad_full) <- 1:nrow(gad_full)

# get_pct <- function(x,y){
#   y <- 1:length(x)
#   for(i in 1:length(x)){if(i==1){y[i] <- round(x[i]/x[i],4)*100} else {y[i] <- round(x[i]/x[1],4)*100}}
#   return(y)
# }

## transfer data to new df and then initialize new column success pct
xgf <- gad_full
xgf$success_pct <- as.numeric(0)
xgf$rel_success_pct <- as.numeric(0)
xgf$hp_success_pct <- as.numeric(0)

head(xgf)
tail(xgf)

#xgf <- xgf[xgf$deviceCategory=="ALL",]

#xgf <- as.data.table(xgf)

#xgf %>% filter(deviceCategory=="tablet" & segment=="00 Homepage" & date==start_date)


## calculate success_pct and relative success_pct per record (using the start_date)
## do this for ALL, desktop, mobile and tablet devices
for (i in 0:n-1){
  ## get the no. of sessions for each step
  ## note that assumption is that xgf contains the columns segment, date and session
  xgf_home <- xgf[xgf$deviceCategory=="ALL" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$sessions
  xgf_slct <- xgf[xgf$deviceCategory=="ALL" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$sessions
  xgf_detl <- xgf[xgf$deviceCategory=="ALL" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$sessions
  xgf_extr <- xgf[xgf$deviceCategory=="ALL" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$sessions
  xgf_pymt <- xgf[xgf$deviceCategory=="ALL" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$sessions
  #xgf_wait <- xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$sessions
  xgf_prch <- xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$sessions
  ## compute for the success pct per step
  success_pct_home <- xgf_prch/xgf_home
  success_pct_slct <- xgf_detl/xgf_slct
  success_pct_detl <- xgf_extr/xgf_detl
  success_pct_extr <- xgf_pymt/xgf_extr
  success_pct_pymt <- xgf_prch/xgf_pymt
  #success_pct_wait <- xgf_prch/xgf_wait
  success_pct_prch <- xgf_prch/xgf_slct
  ## compute for the relative success pct per step
  r_success_pct_home <- xgf_prch/xgf_home
  r_success_pct_slct <- xgf_slct/xgf_slct
  r_success_pct_detl <- xgf_detl/xgf_slct
  r_success_pct_extr <- xgf_extr/xgf_slct
  r_success_pct_pymt <- xgf_pymt/xgf_slct
  #r_success_pct_wait <- xgf_wait/xgf_slct
  r_success_pct_prch <- xgf_prch/xgf_slct
  ## success pct vs homepage
  r1_success_pct_home <- xgf_prch/xgf_home
  r1_success_pct_slct <- xgf_slct/xgf_home
  r1_success_pct_detl <- xgf_detl/xgf_home
  r1_success_pct_extr <- xgf_extr/xgf_home
  r1_success_pct_pymt <- xgf_pymt/xgf_home
  #r1_success_pct_wait <- xgf_wait/xgf_home
  r1_success_pct_prch <- xgf_prch/xgf_home
  ## assign the values per step per different variable and add a new column for the success pct 
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$success_pct <- success_pct_home
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$success_pct <- success_pct_slct
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$success_pct <- success_pct_detl
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$success_pct <- success_pct_extr
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$success_pct <- success_pct_pymt
  #xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$success_pct <- success_pct_wait
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$success_pct <- success_pct_prch
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_home
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_slct
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_detl
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_extr
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_pymt
  #xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_wait
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_prch
  ## assign the values per step for the success pct vs homepage
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_home
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_slct
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_detl
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_extr
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_pymt
  #xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_wait
  xgf[xgf$deviceCategory=="ALL" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_prch
  ## increment dates
  start_date <- start_date + 1
}

head(xgf)
tail(xgf)

#xgf[xgf$date=="2019-07-21",]

n <- 1
## re-initialize dates
end_date <- Sys.Date() - 1 # current date -1 to ensure that complete data will be fetched
end_date <- as.Date(end_date) # ensuring date format
start_date <- end_date 
start_date <- as.Date(start_date)
start_date
end_date


## calculate success_pct and relative success_pct per record (using the start_date)
## do this for ALL, desktop, mobile and tablet devices
for (i in 0:n-1){
  ## get the no. of sessions for each step
  ## note that assumption is that xgf contains the columns segment, date and session
  xgf_home <- xgf[xgf$deviceCategory=="desktop" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$sessions
  xgf_slct <- xgf[xgf$deviceCategory=="desktop" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$sessions
  xgf_detl <- xgf[xgf$deviceCategory=="desktop" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$sessions
  xgf_extr <- xgf[xgf$deviceCategory=="desktop" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$sessions
  xgf_pymt <- xgf[xgf$deviceCategory=="desktop" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$sessions
  #xgf_wait <- xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$sessions
  xgf_prch <- xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$sessions
  ## compute for the success pct per step
  success_pct_home <- xgf_prch/xgf_home
  success_pct_slct <- xgf_detl/xgf_slct
  success_pct_detl <- xgf_extr/xgf_detl
  success_pct_extr <- xgf_pymt/xgf_extr
  success_pct_pymt <- xgf_prch/xgf_pymt
  #success_pct_wait <- xgf_prch/xgf_wait
  success_pct_prch <- xgf_prch/xgf_slct
  ## compute for the relative success pct per step
  r_success_pct_home <- xgf_prch/xgf_home
  r_success_pct_slct <- xgf_slct/xgf_slct
  r_success_pct_detl <- xgf_detl/xgf_slct
  r_success_pct_extr <- xgf_extr/xgf_slct
  r_success_pct_pymt <- xgf_pymt/xgf_slct
  #r_success_pct_wait <- xgf_wait/xgf_slct
  r_success_pct_prch <- xgf_prch/xgf_slct
  ## success pct vs homepage
  r1_success_pct_home <- xgf_prch/xgf_home
  r1_success_pct_slct <- xgf_slct/xgf_home
  r1_success_pct_detl <- xgf_detl/xgf_home
  r1_success_pct_extr <- xgf_extr/xgf_home
  r1_success_pct_pymt <- xgf_pymt/xgf_home
  #r1_success_pct_wait <- xgf_wait/xgf_home
  r1_success_pct_prch <- xgf_prch/xgf_home
  ## assign the values per step per different variable and add a new column for the success pct 
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$success_pct <- success_pct_home
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$success_pct <- success_pct_slct
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$success_pct <- success_pct_detl
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$success_pct <- success_pct_extr
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$success_pct <- success_pct_pymt
  #xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$success_pct <- success_pct_wait
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$success_pct <- success_pct_prch
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_home
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_slct
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_detl
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_extr
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_pymt
  #xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_wait
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_prch
  ## assign the values per step for the success pct vs homepage
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_home
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_slct
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_detl
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_extr
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_pymt
  #xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_wait
  xgf[xgf$deviceCategory=="desktop" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_prch
  ## increment dates
  start_date <- start_date + 1
}

head(xgf)

n <- 1
## re-initialize dates
end_date <- Sys.Date() - 1 # current date -1 to ensure that complete data will be fetched
end_date <- as.Date(end_date) # ensuring date format
start_date <- end_date 
start_date <- as.Date(start_date)
start_date
end_date


## calculate success_pct and relative success_pct per record (using the start_date)
## do this for ALL, desktop, mobile and tablet devices
for (i in 0:n-1){
  ## get the no. of sessions for each step
  ## note that assumption is that xgf contains the columns segment, date and session
  xgf_home <- xgf[xgf$deviceCategory=="mobile" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$sessions
  xgf_slct <- xgf[xgf$deviceCategory=="mobile" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$sessions
  xgf_detl <- xgf[xgf$deviceCategory=="mobile" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$sessions
  xgf_extr <- xgf[xgf$deviceCategory=="mobile" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$sessions
  xgf_pymt <- xgf[xgf$deviceCategory=="mobile" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$sessions
  #xgf_wait <- xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$sessions
  xgf_prch <- xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$sessions
  ## compute for the success pct per step
  success_pct_home <- xgf_prch/xgf_home
  success_pct_slct <- xgf_detl/xgf_slct
  success_pct_detl <- xgf_extr/xgf_detl
  success_pct_extr <- xgf_pymt/xgf_extr
  success_pct_pymt <- xgf_prch/xgf_pymt
  #success_pct_wait <- xgf_prch/xgf_wait
  success_pct_prch <- xgf_prch/xgf_slct
  ## compute for the relative success pct per step
  r_success_pct_home <- xgf_prch/xgf_home
  r_success_pct_slct <- xgf_slct/xgf_slct
  r_success_pct_detl <- xgf_detl/xgf_slct
  r_success_pct_extr <- xgf_extr/xgf_slct
  r_success_pct_pymt <- xgf_pymt/xgf_slct
  #r_success_pct_wait <- xgf_wait/xgf_slct
  r_success_pct_prch <- xgf_prch/xgf_slct
  ## success pct vs homepage
  r1_success_pct_home <- xgf_prch/xgf_home
  r1_success_pct_slct <- xgf_slct/xgf_home
  r1_success_pct_detl <- xgf_detl/xgf_home
  r1_success_pct_extr <- xgf_extr/xgf_home
  r1_success_pct_pymt <- xgf_pymt/xgf_home
  #r1_success_pct_wait <- xgf_wait/xgf_home
  r1_success_pct_prch <- xgf_prch/xgf_home
  ## assign the values per step per different variable and add a new column for the success pct 
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$success_pct <- success_pct_home
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$success_pct <- success_pct_slct
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$success_pct <- success_pct_detl
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$success_pct <- success_pct_extr
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$success_pct <- success_pct_pymt
  #xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$success_pct <- success_pct_wait
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$success_pct <- success_pct_prch
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_home
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_slct
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_detl
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_extr
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_pymt
  #xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_wait
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_prch
  ## assign the values per step for the success pct vs homepage
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_home
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_slct
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_detl
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_extr
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_pymt
  #xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_wait
  xgf[xgf$deviceCategory=="mobile" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_prch
  ## increment dates
  start_date <- start_date + 1
}



n <- 1
## re-initialize dates
end_date <- Sys.Date() - 1 # current date -1 to ensure that complete data will be fetched
end_date <- as.Date(end_date) # ensuring date format
start_date <- end_date 
start_date <- as.Date(start_date)
start_date
end_date


## calculate success_pct and relative success_pct per record (using the start_date)
## do this for ALL, desktop, mobile and tablet devices
for (i in 0:n-1){
  ## get the no. of sessions for each step
  ## note that assumption is that xgf contains the columns segment, date and session
  #start_date <- "2020-05-24"
  
  #xgf <- as.data.table(xgf)
  
  xgf_home <- xgf[xgf$deviceCategory=="tablet" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$sessions
  xgf_slct <- xgf[xgf$deviceCategory=="tablet" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$sessions
  xgf_detl <- xgf[xgf$deviceCategory=="tablet" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$sessions
  xgf_extr <- xgf[xgf$deviceCategory=="tablet" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$sessions
  xgf_pymt <- xgf[xgf$deviceCategory=="tablet" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$sessions
  #xgf_wait <- xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$sessions
  xgf_prch <- xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$sessions
  ## compute for the success pct per step
  success_pct_home <- xgf_prch/xgf_home
  success_pct_slct <- xgf_detl/xgf_slct
  success_pct_detl <- xgf_extr/xgf_detl
  success_pct_extr <- xgf_pymt/xgf_extr
  success_pct_pymt <- xgf_prch/xgf_pymt
  #success_pct_wait <- xgf_prch/xgf_wait
  success_pct_prch <- xgf_prch/xgf_slct
  ## compute for the relative success pct per step
  r_success_pct_home <- xgf_prch/xgf_home
  r_success_pct_slct <- xgf_slct/xgf_slct
  r_success_pct_detl <- xgf_detl/xgf_slct
  r_success_pct_extr <- xgf_extr/xgf_slct
  r_success_pct_pymt <- xgf_pymt/xgf_slct
  #r_success_pct_wait <- xgf_wait/xgf_slct
  r_success_pct_prch <- xgf_prch/xgf_slct
  ## success pct vs homepage
  r1_success_pct_home <- xgf_prch/xgf_home
  r1_success_pct_slct <- xgf_slct/xgf_home
  r1_success_pct_detl <- xgf_detl/xgf_home
  r1_success_pct_extr <- xgf_extr/xgf_home
  r1_success_pct_pymt <- xgf_pymt/xgf_home
  #r1_success_pct_wait <- xgf_wait/xgf_home
  r1_success_pct_prch <- xgf_prch/xgf_home
  ## assign the values per step per different variable and add a new column for the success pct 
  
  if (nrow(xgf[xgf$deviceCategory=="tablet" & xgf$segment=="00 Homepage" & xgf$date==start_date, ])>0){
    xgf[xgf$deviceCategory=="tablet" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$success_pct <- success_pct_home
  }
  
  if (nrow(xgf[xgf$deviceCategory=="tablet" & xgf$segment=="01 New - Select" & xgf$date==start_date, ])>0){
    xgf[xgf$deviceCategory=="tablet" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$success_pct <- success_pct_slct
  }
  
  if (nrow(xgf[xgf$deviceCategory=="tablet" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ])>0){
    xgf[xgf$deviceCategory=="tablet" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$success_pct <- success_pct_detl
  }
  
  if (nrow(xgf[xgf$deviceCategory=="tablet" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ])>0){
    xgf[xgf$deviceCategory=="tablet" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$success_pct <- success_pct_extr
  }
  
  if (nrow(xgf[xgf$deviceCategory=="tablet" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ])>0){
    xgf[xgf$deviceCategory=="tablet" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$success_pct <- success_pct_pymt
  }
  
  #if (nrow(xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ])>0){
  #  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$success_pct <- success_pct_wait
  #}
  
  if (nrow(xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ])>0){
    xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$success_pct <- success_pct_prch
  }
  
  
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_home
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_slct
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_detl
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_extr
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_pymt
  #xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_wait
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$rel_success_pct <- r_success_pct_prch
  ## assign the values per step for the success pct vs homepage
  ## assign the values per step for the relative success pct
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="00 Homepage" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_home
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="01 New - Select" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_slct
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="02 New - Guest Details" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_detl
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="03 New - Extras" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_extr
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="04 New - Payment" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_pymt
  #xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Wait" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_wait
  xgf[xgf$deviceCategory=="tablet" & xgf$segment=="05 New - Purchase Ticket" & xgf$date==start_date, ]$hp_success_pct <- r1_success_pct_prch
  ## increment dates
  start_date <- start_date + 1
}

head(xgf,20)
tail(xgf)

n <- 1
## re-initialize dates
end_date <- Sys.Date() - 1 # current date -1 to ensure that complete data will be fetched
end_date <- as.Date(end_date) # ensuring date format
start_date <- end_date 
start_date <- as.Date(start_date)
start_date
end_date

gad_full_calc <- xgf
head(gad_full_calc)

## prepare data for reporting (formatting)
#dummy_df <- data.frame("dummy")
gad_full_calc$rel_success_pct <- round(gad_full_calc$rel_success_pct, 4)
#gadf_per_period$rel_success_pct <- round(gadf_per_period$rel_success_pct, 4)
rept_gad_full_calc <- gad_full_calc
#rept_gadf_per_period <- gadf_per_period
head(rept_gad_full_calc)
#head(rept_gadf_per_period)
#rept_gadf_per_period$date <- rept_gadf_per_period$period
#rept_gadf_per_period$period <- NULL
#rept_gadf_per_period$deviceCategory <- "Overall"
rept_gad_full_calc$date <- as.character.Date(rept_gad_full_calc$date)
#full_rept_ga_df <- rbind(rept_gad_full_calc, rept_gadf_per_period)
full_rept_ga_df <- rept_gad_full_calc
tail(full_rept_ga_df)
head(full_rept_ga_df)
str(full_rept_ga_df)
full_rept_ga_df <- data.frame(full_rept_ga_df, stringsAsFactors = FALSE)
full_rept_ga_df$segment <- as.character(full_rept_ga_df$segment)
full_rept_ga_df$deviceCategory[full_rept_ga_df$deviceCategory=="ALL"] <- "Overall"
full_rept_ga_df$deviceCategory[full_rept_ga_df$deviceCategory=="desktop"] <- "Desktop"
full_rept_ga_df$deviceCategory[full_rept_ga_df$deviceCategory=="mobile"] <- "Mobile"
full_rept_ga_df$deviceCategory[full_rept_ga_df$deviceCategory=="tablet"] <- "Tablet"
full_rept_ga_df$segment[full_rept_ga_df$segment=="00 Homepage"]  <- "00 Homepage"
full_rept_ga_df$segment[full_rept_ga_df$segment=="01 New - Select"]  <- "01 Select"
full_rept_ga_df$segment[full_rept_ga_df$segment=="02 New - Guest Details"]  <- "02 Guest Details"
full_rept_ga_df$segment[full_rept_ga_df$segment=="03 New - Extras"]  <- "03 Add-Ons"
full_rept_ga_df$segment[full_rept_ga_df$segment=="04 New - Payment"]  <- "04 Payment"
#full_rept_ga_df$segment[full_rept_ga_df$segment=="05 New - Wait"]  <- "05 Wait"
full_rept_ga_df$segment[full_rept_ga_df$segment=="05 New - Purchase Ticket"]  <- "05 Purchase"
full_rept_ga_df$segment <- factor(full_rept_ga_df$segment)
full_rept_ga_df$date <- factor(full_rept_ga_df$date)
full_rept_ga_df$deviceCategory <- factor(full_rept_ga_df$deviceCategory)

head(full_rept_ga_df)
tail(full_rept_ga_df)



fnl_file <- "OMNIX Device Funnel.xlsx"
file_dir <- "C:/Users/AnalyticsReporting/Desktop/ARCHIE/Marketing/MK_FNL/DAILY/DATA"
prev_fnl <- readWorkbook(paste0(file_dir,"/",fnl_file), sheet='OMNIX 1D - Raw Data', detectDates=T)
prev_fnl <- prev_fnl %>% arrange(desc(date),deviceCategory, segment)

prev_fnl$latestdate <- if_else(prev_fnl$latestdate == Sys.Date() - 1, "TRUE", "FALSE")

head(prev_fnl)



full_rept_ga_df <- as.data.frame(full_rept_ga_df)


full_rept_ga_df$date <- as.Date(full_rept_ga_df$date)

depr_df <- full_rept_ga_df %>% 
  mutate(latestdate = if_else(full_rept_ga_df$date == Sys.Date() - 1, "TRUE", "FALSE"))


fnl_all <- rbind(depr_df, prev_fnl)
fnl_all <- fnl_all %>% arrange(desc(date), deviceCategory, segment)

head(fnl_all)
tail(fnl_all)



wb1 <- loadWorkbook("C:/Users/AnalyticsReporting/Desktop/ARCHIE/Marketing/MK_FNL/DAILY/DATA/OMNIX Device Funnel.xlsx")
writeData(wb1, sheet = "OMNIX 1D - Raw Data", fnl_all)
saveWorkbook(wb1,"C:/Users/AnalyticsReporting/Desktop/ARCHIE/Marketing/MK_FNL/DAILY/DATA/OMNIX Device Funnel.xlsx",overwrite = T)



wd <- "C:/Users/AnalyticsReporting/Desktop/ARCHIE/Marketing/MK_FNL/DAILY/DATA/"

OutApp <- COMCreate("Excel.Application")
filename <- paste0(wd,'OMNIX Device Funnel.xlsx')

PPT <- OutApp$Workbooks()$Open(filename)
PPT$RefreshAll()
Sys.sleep(10)
PPT$Saved()
PPT$Save()
PPT$Saved()
PPT$Close()


### UPDATE POWERPOINT LINK ####
OutApp <- COMCreate("Powerpoint.Application")
filename <- 'C:\\Users\\AnalyticsReporting\\Desktop\\ARCHIE\\Marketing\\MK_FNL\\DAILY\\DATA\\OMNIX Daily.pptx'
PPT <- OutApp$Presentations()$Open(filename)
PPT$UpdateLinks()
Sys.sleep(5)
PPT$Saved()
PPT$Save()
PPT$Saved()


path <- 'C:\\Users\\AnalyticsReporting\\Desktop\\ARCHIE\\Marketing\\MK_FNL\\DAILY\\DATA\\Notif_image1.jpg'
PPT$Slides(1)$Export(path,"JPG")
PPT$Saved()
PPT$Save()
PPT$Saved()

PPT$Close()


timestamp <- format(Sys.Date()-1,"%Y%m%d")
asof <- format(Sys.Date()-1,"%Y-%m-%d")

####################### SHAREPOINT UPLOAD
spwd <- paste0("T:\\Digital Performance\\Daily Funnel Reports\\") ####make sharepoint folder
dir.create(spwd,recursive=TRUE)

orig <- paste0(wd,fnl_file)
copy <- paste0(spwd,paste0("OMNIX Device Funnel_",timestamp,".xlsx"))
file.copy(orig,copy,overwrite = TRUE)

report_link <- paste0("https://iamcebv3.cebupacificair.com/departments/analytics/Analytics%20Reports/Digital%20Performance/Daily%20Funnel%20Reports/OMNIX%20Device%20Funnel_",timestamp,".xlsx")



######################### UPDATE ANALYTICSSANDBOX

library(readxl)
mypath = "T:/Digital Performance/Daily Funnel Reports"
setwd(mypath)

yday_date <- format(Sys.Date() - 1,"%Y%m%d")

my_data <- read_excel(paste0("OMNIX Device Funnel_",yday_date,".xlsx"), sheet = 2)

fnl_sessions <- my_data %>% filter(latestdate=="TRUE") %>% select(date, deviceCategory, segment, sessions)


fnl_sessions$date <- as.Date(fnl_sessions$date)

colnames(fnl_sessions) <- c("fnl_date", "devicecategory", "segment", "sessions")

fnl_pct_calc <- fnl_sessions %>% arrange(desc(fnl_date), devicecategory, segment) %>%
  group_by(fnl_date,devicecategory) %>%
  mutate(pct_homepage = sessions/sessions[segment == '00 Homepage'], pct_select = sessions/sessions[segment == '01 Select'], pct_prior = sessions/lag(sessions))

pct_df <- fnl_pct_calc

is.na(pct_df$pct_select) <- pct_df$segment == "00 Homepage"


############### PG CONNECTIONS
#pg <- dbDriver("PostgreSQL")
#con <- dbConnect(pg, user="jrjoanna", password="o*?XJCG$@Q#A",
                # host="10.80.137.11", port=5432, dbname="daylight")


#dbWriteTable(con,c('analyticssandbox', 'ga_mkfnl_da_omnixfnldata'),pct_df,
         #    row.names=FALSE,overwrite=FALSE, append=TRUE)
#dbSendQuery(con,'GRANT ALL ON TABLE analyticssandbox.ga_mkfnl_da_omnixfnldata TO analytics;')





######################### UPDATE OMNIXFUNNEL CSV FILE IN SHAREPOINT


FILE_NAME <- "OmnixFunnel.csv"

input_data <- read.csv(file = FILE_NAME, header = TRUE, sep = ',', stringsAsFactors = FALSE)
result <- data.frame(input_data, stringsAsFactors = FALSE)

fnl_pbi <- data.frame(pct_df, stringsAsFactors = FALSE)
pbi_update <- rbind(fnl_pbi,result)

pbi_update[is.na(pbi_update)] <- ""

write.csv(pbi_update, "OmnixFunnel.csv", row.names = FALSE, quote = FALSE)



######################### SEND EMAIL NOTIFICATION


setwd("C:\\Users\\AnalyticsReporting\\Desktop\\Source Codes\\r\\emailfunctions")
source("sendemail v2.R")



################################### EMAIL BODY #################################
subject1<- paste0("OMNIX Device Funnel Report as of ",asof)
taskname <- 'MK_FNL_01'


screenshot <- "C:\\Users\\AnalyticsReporting\\Desktop\\ARCHIE\\Marketing\\MK_FNL\\DAILY\\DATA\\Notif_image1.jpg"
OutMail[["Attachments"]]$Add(screenshot)
cid <- 'Not this'
i <- 0
while (cid != 'Notif_image1.jpg'){
  i = i+1
  cid <- OutMail[["Attachments"]]$Item(i)$FileName()
}
OutMail[["Attachments"]]$Item(i)$PropertyAccessor()$SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F","Notif_image1.jpg")
screenshot <- '<img src="Notif_image1.jpg" width = "800">'


{
  body1 <-paste0(
    "Hi Everyone, <br><br>The OMNIX Device Funnel Report as of <b>",
    asof,"</b> has been generated and is now available for download. Please click this <a href=",report_link,">link</a> to download the report.<br><br>", screenshot, "<br><br> For questions/concerns, feel free to hit the <b>reply or reply all button</b> to let the assigned analyst know.<br>"
  )
  send_email(taskname,subject1,body1,0,0)
}
