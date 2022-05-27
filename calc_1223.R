#-----------------install packages that will be used--------------------#
my_packages<-c("tidyverse","lubridate","readr","rio","readxl",
               "janitor","purrr","data.table","data.table","mefa",
               "openxlsx", "ggplot2")

new_packages<-my_packages[!(my_packages %in% utils::installed.packages())]

if(!require(new_packages)) {
  install.packages(new_packages, dependencies = TRUE)}

#----------------------------load packages-------------------------------#
library(tidyverse)
library(lubridate)
library(readr)
library(rio)
library(readxl)
library(janitor)
library(purrr)
library(data.table)
library(mefa)
library(openxlsx)
library(ggplot2)
#----------set memory/local language for R running, optional-------------#
memory.limit(size = 25000)
Sys.setlocale(locale = "chs")
#------find a location and create one folder for running-----------------#
#--------------define your directory-------------------------------------#
my_dir<-"C:/abcd"
#-----------------------------BACK UP -------------------#
print("create folders for decomped results and back up last mdoel results")
if(!file.exists(paste(my_dir,"results", sep="/"))){
  dir.create(path = paste(my_dir,"results", sep="/"))
}
if(!file.exists(paste(my_dir,"backup", sep="/"))){
  dir.create(path = paste(my_dir,"backup", sep="/"))
}

print("starting back up") 

oldfiles <- list.files(paste(my_dir,"results", sep="/"), full.names = FALSE)
dir.create(path = paste(my_dir,"backup", format(Sys.time(), "%Y-%m-%d %H-%M"),sep="/"))

lapply(oldfiles, function(x) file.copy(paste(my_dir,"results", x , sep = "/"),  
                                       paste (my_dir,"backup", format(Sys.time(), "%Y-%m-%d %H-%M"), x, sep = "/"), 
                                       recursive = FALSE,  copy.mode = TRUE))

print("back up finished")
print("starting generate calc---")

#--------------------------read files and get sales-----------------------#
#-----------------------------week map------------------------------------#
#-------please make sure you have these files under your directory--------#
#--------time map.xlsx, spending, your decomp files prepared--------------#

week_map<-read_excel(paste(my_dir,"time map.xlsx", sep="/"),
                     sheet = "week")%>%
  mutate(week=ymd(week))

spending<-read_excel(paste(my_dir,"spending.xlsx", sep="/"),
                     sheet = "raw")%>%
  clean_names()%>%
  mutate(week=ymd(week))
spending<-setDT(spending)
setnames(spending, "var_1", "var_2")
spending<-spending[,by=.(mkt, variable_key, var_map, 
                         var_chn1, var_chn2, var_chn3, var_chn4, week,metric_name),
                   .(actual_spending_rmb=sum(actual_spending_rmb, na.rm=TRUE),
                     metric=sum(metric,na.rm=TRUE),
                     clicks=sum(clicks,na.rm=TRUE),
                     impressions=sum(impressions,na.rm=TRUE),
                     engagement=sum(engagement,na.rm=TRUE),
                     pr_value_rmb=sum(pr_value_rmb,na.rm=TRUE),
                     views=sum(views,na.rm=TRUE))]

#--------------define functions------------------------#
sales_decompose<-function(file_name){
  raw_decomp<-read_excel(file_name, sheet = "decomp")
  raw_decomp<-setDT(raw_decomp)
  raw_decomp<-raw_decomp[,week:=mdy(week)]
  my_cols<-names(raw_decomp)[!(grepl(pattern = "*_CUR", names(raw_decomp))|
                                 grepl(pattern = "*_PRV", names(raw_decomp)))]
  raw_decomp<-raw_decomp[,..my_cols]
  raw_decomp<-remove_empty(raw_decomp, which = "cols")
  raw_decomp<-raw_decomp[Geogname!="TOTGC",]
  id_list<-c("IPweight", "prodkey", "Prodname","Prodlvl","geogkey",
             "Geogname","Geoglvl","week","BRAND")
  raw_decomp<-melt.data.table(raw_decomp, id.vars = id_list)
  setnames(raw_decomp, "value", "decomp_unit")
  
  product_map<-read_excel(file_name, sheet = "Product map")
  product_map<-setDT(product_map)
  variable_map<-read_excel(file_name, sheet = "var_mapping")
  variable_map<-setDT(variable_map)
  raw_decomp<-merge(raw_decomp, variable_map, by.x="variable", by.y="var",
                    all.x=TRUE)
  raw_decomp<-raw_decomp[var_map!="Exclude",]
  
  revenue<-read_excel(file_name, sheet = "revenue")
  revenue<-setDT(revenue)
  revenue<-revenue[,week:=ymd(week)]
  revenue<-revenue[,c("channel","Geogname", "season", "week", "sold_price","net_price","bumpup_factor")]
  raw_decomp<-merge(raw_decomp, product_map, by.x=c("prodkey","Geogname"), 
                    by.y=c("prodkey","Geogname"), all.x=TRUE)
  raw_decomp<-merge(raw_decomp,week_map, by.x="week", by.y="week", all.x=TRUE)
  raw_decomp<-merge(raw_decomp, revenue, by.x=c("channel","season","week","Geogname"),
                    by.y=c("channel","season","week","Geogname"),all.x=TRUE)
  
  raw_decomp<-raw_decomp[grepl("Sales Unit",product)
                         ,decomp_unit_bump:=decomp_unit*bumpup_factor]
  raw_decomp<-raw_decomp[grepl("Traffic",product)
                         ,decomp_unit_bump:=decomp_unit]
  
  raw_decomp<-raw_decomp[grepl("Sales Unit",product)
                         ,c("decomp_value","decomp_NetRevenue"):=list(decomp_unit_bump*sold_price,
                                                                      decomp_unit_bump*net_price)]
  raw_decomp<-raw_decomp[grepl("Traffic",product)
                         ,c("decomp_value","decomp_NetRevenue"):=list(decomp_unit,
                                decomp_unit)]
  
  return(raw_decomp)
  
}

#----------------redistributed results---------------------------#
# my_raw_decomp<-tmall_ym
#-----------------traffic decomp---------------------------------#
re_distribute<-function(my_raw_decomp){
  traffic_decomp<-my_raw_decomp[grepl("Traffic",product),]
  sales_decomp<-my_raw_decomp[grepl("Sales Unit",product),]
  
  traffic_actual<-traffic_decomp[var_map=="Actual",
                                 by=.(channel, week, product, Geogname, Prodname, prodkey,
                                      Prodlvl, geogkey, Geoglvl, BRAND, variable),
                                 .(actual_weekly=sum(decomp_unit_bump, na.rm=TRUE))]
  
  traffic_decomp<-merge(traffic_decomp,traffic_actual,
                        by=c("channel", "week", "product", "Geogname", "Prodname", "prodkey",
                             "Prodlvl", "geogkey", "Geoglvl", "BRAND"),
                        all.x=TRUE)
  traffic_decomp<-traffic_decomp[,variable.y:=NULL]
  setnames(traffic_decomp, "variable.x", "variable")
  traffic_decomp<-traffic_decomp[,traffic_contri:=decomp_unit_bump/actual_weekly]
  
  sales_decomp_traffic<-sales_decomp[var_map=="Traffic",
                                     by=.(channel, week, product, Geogname, Prodname, prodkey,
                                          Prodlvl, geogkey, Geoglvl, BRAND, variable),
                                     .(traffic_weekly=sum(decomp_unit_bump, na.rm=TRUE))]
  
  sales_decomp_traffic<-merge(traffic_decomp,
                              sales_decomp_traffic[,-c("variable","product",
                                                       "Prodname","prodkey","BRAND")],
                              by=c("channel", "week", 
                                   "Geogname",
                                   "Prodlvl", "geogkey", "Geoglvl"), 
                              all.x=TRUE)
  sales_decomp_traffic<-sales_decomp_traffic[,traffic_unit:=traffic_weekly*traffic_contri]
  
  # sales_decomp_traffic<-sales_decomp_traffic[!(var_map %in%
  #                               c("Intercept","Actual", "Predict", "Residual")),]
  sales_traffic_mkt<-sales_decomp_traffic[!(var_map %in%
                                              c("Intercept","Actual", "Predict", "Residual",
                                                "Temperature","Fourier_season","Seasonality","Store_Count",
                                                "Baidu index")),
                                          by=.(channel, week, Geogname,  
                                               Prodlvl, geogkey, Geoglvl),
                                          .(traffic_mkt=sum(traffic_unit, na.rm=TRUE))]
  
  sales_decomp_traffic<-merge(sales_decomp_traffic, sales_traffic_mkt, 
                              by=c("channel", "week", "Geogname",  
                                   "Prodlvl", "geogkey", "Geoglvl"))
  my_cols<-c("channel", "week", "Geogname",
             "Prodlvl", "geogkey", "Geoglvl",
             "traffic_contri", "variable",
             "actual_weekly","traffic_weekly", "traffic_unit","traffic_mkt")
  
  sales_decomp<-merge(sales_decomp, 
                      sales_decomp_traffic[!(var_map %in%
                                               c("Intercept","Actual", "Predict", "Residual",
                                                 "Temperature","Fourier_season","Seasonality","Store_Count",
                                                 "Baidu index")),..my_cols],
                      by=c("channel", "week", "Geogname",
                           "Prodlvl", "geogkey", "Geoglvl",
                           "variable"),
                      all=TRUE)                         
  sales_decomp<-setnafill(sales_decomp, fill = 0, cols = c("traffic_unit","traffic_mkt"))
  sales_decomp<-sales_decomp[,decomp_unit_adjust:= decomp_unit_bump + traffic_unit]
  sales_decomp<-sales_decomp[var_map=="Traffic",
                             decomp_unit_adjust:= decomp_unit_bump - traffic_mkt]
  
  sales_decomp<-sales_decomp[,c("decomp_value_adjust", "decomp_NetRevenue_adjust"):= 
                               list(decomp_unit_adjust*sold_price,
                                    decomp_unit_adjust*net_price )]
  traffic_decomp<-traffic_decomp[,c("decomp_unit_adjust", "decomp_value_adjust","decomp_NetRevenue_adjust"):=
                                   list(decomp_unit, decomp_value, decomp_value)]
  redistributed_decomp<-bind_rows(sales_decomp, traffic_decomp)
  return(redistributed_decomp)
  
}
#--------------------------map spending----------------------------------#

map_spd<-function(my_dataframe, spd_data){
  key_names<-c("var_chn1", "var_chn2", "var_chn3", "var_chn4")
  my_key<-key_names[grepl(substitute(my_dataframe),key_names)]
  my_cols<-c("mkt","week",my_key,"actual_spending_rmb",
             "metric","impressions","clicks")
  spd<-spd_data[,..my_cols]
  setnames(spd, c(my_key), c("var_key"))
  spd<-spd[,lapply(.SD, sum, na.rm=TRUE),
           by=.(mkt, var_key, week),
           .SDcols=c("actual_spending_rmb",
                     "metric", "impressions","clicks")]
  my_dataframe<-setDT(my_dataframe)[,by=.(channel,week, Geogname, 
                                          Prodlvl, geogkey, Geoglvl,season,
                                          prodkey, Prodname, BRAND, var_map,var_key,
                                          group2, group3, group4, product, year, mat),
                                    .(decomp_unit=sum(decomp_unit, na.rm=TRUE),
                                      decomp_value=sum(decomp_value, na.rm=TRUE),
                                      decomp_unit_adjust=sum(decomp_unit_adjust, na.rm=TRUE),
                                      decomp_value_adjust=sum(decomp_value_adjust, na.rm=TRUE),
                                      sold_price=mean(sold_price, na.rm=TRUE))]
  my_dataframe<-merge(my_dataframe, spd, by.x=c("group4", "var_key", "week"),
                      by.y=c("mkt", "var_key", "week"),
                      all.x=TRUE)
  return(my_dataframe)
}


#--------------------------------decompse-------------------------#
decomp_file<-list.files(paste(my_dir, "decomp", sep="/"), 
                        full.names = TRUE, pattern = "*xlsx")


data_1<-sales_decompose(decomp_file[grepl(pattern = "pattern1",decomp_file)])
data_1<-re_distribute(my_raw_decomp=data_1)

data_2<-sales_decompose(decomp_file[grepl(pattern = "pattern2",decomp_file)])
data_2<-re_distribute(my_raw_decomp=data_2)

data_3<-sales_decompose(decomp_file[grepl(pattern = "pattern3",decomp_file)])
data_3<-re_distribute(my_raw_decomp=whs)

data_4<-sales_decompose(decomp_file[grepl(pattern = "pattern4",decomp_file)])
data_4<-re_distribute(my_raw_decomp=data_4)

my_decomp<-bind_rows(data_1,data_2,data_3,data_4)
my_decomp<-my_decomp[, metric:= "Sales Unit"]
my_decomp<-my_decomp[grepl("Traffic",Prodname),metric:= "Traffic"]

###aggregate
my_decomp_agg<-my_decomp[,by=.(metric, mat, season, 
                                 channel, Geogname,
                                 Prodlvl, Geoglvl,
                                 variable, IPweight,
                                 var_map, group2, group3,
                                 group4, var_key, year, 
                                 Prodname,
                                 geogkey, prodkey,
                                 BRAND, product),
                           .(decomp_unit_bump=sum(decomp_unit_bump,na.rm=TRUE), 
                             net_price=mean(net_price, na.rm=TRUE),
                             decomp_NetRevenue=sum(decomp_NetRevenue, na.rm=TRUE),
                             decomp_unit_adjust=sum(decomp_unit_adjust, na.rm=TRUE),
                             decomp_NetRevenue_adjust =sum(decomp_NetRevenue_adjust, na.rm = TRUE),
                             traffic_unit=sum(traffic_unit, na.rm=TRUE),
                             traffic_mkt =sum(traffic_mkt, na.rm=TRUE),
                             sold_price=mean(sold_price, na.rm=TRUE),
                             decomp_value=sum(decomp_value, na.rm=TRUE),
                             decomp_value_adjust =sum(decomp_value_adjust, na.rm = TRUE),
                             decomp_unit=sum(decomp_unit,na.rm=TRUE),
                             bumpup_factor=mean(bumpup_factor, na.rm=TRUE))]


#------------------------model fit--------------------------#

my_decomp_fit<-my_decomp[variable=="_Actual"|variable=="_FinlFit", 
                           by=.(channel, Geogname,product, 
                                week, season, year, mat, var_map, metric),
                           .(unitbump=sum(decomp_unit_bump, na.rm=TRUE),
                             netvalue=sum(decomp_NetRevenue,na.rm=TRUE),
                            grossvalue=sum(decomp_value,na.rm=TRUE),
                            unit=sum(decomp_unit, na.rm=TRUE))]%>%
  pivot_wider(names_from = var_map, values_from = c(unitbump, netvalue,grossvalue, unit))%>%
  arrange(channel,metric , product, Geogname, week)

my_decomp_fit<-setDT(my_decomp_fit)

#------------------------------------write ----------------------------------------------#
setcolorder(my_decomp, neworder = c("week","Prodlvl","Geoglvl","variable","season",
                                   "IPweight", "var_map" , "group2" ,"group3",
                                   "group4", "var_key", "year", "mat", "metric",
                                   "decomp_unit_bump", "net_price", "decomp_NetRevenue", 
                                   "decomp_unit_adjust",
                                   "decomp_NetRevenue_adjust","traffic_unit",
                                   "traffic_mkt","channel", "Geogname", "Prodname",
                                   "geogkey","prodkey", "BRAND", "product",
                                   "traffic_contri","actual_weekly","traffic_weekly",
                                   "sold_price" ,"decomp_value" ,"decomp_value_adjust",
                                   "decomp_unit","bumpup_factor"))



# Sys.setlocale(locale = "chs")
print("!calc generating finished")



#-------------------write csv， the other time --------------------------------#
write_csv(my_decomp, paste(my_dir,"results", "my_baseline_decomp.csv", sep="/"))
write_csv(my_decomp_fit, paste(my_dir,"results", "my_actual_predict.csv", sep="/"))
write_csv(my_decomp_agg, paste(my_dir,"results", "my_actual_agg.csv", sep="/"))


print("!calc generating finished")

