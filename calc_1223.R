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
my_dir<-"C:/005 Decker/005 Decker/002 model/calc"
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

spending<-read_excel(paste(my_dir,"spending ugg.xlsx", sep="/"),
                     sheet = "raw")%>%
  clean_names()%>%
  mutate(week=ymd(week))
spending<-setDT(spending)
setnames(spending, "var_ugg_cn", "var_ugg.cn")
spending<-spending[,by=.(mkt, variable_key, var_map, 
                         var_ugg.cn, var_tmall, var_whs, var_dtc, week,metric_name),
                   .(actual_spending_rmb=sum(actual_spending_rmb, na.rm=TRUE),
                     metric=sum(metric,na.rm=TRUE),
                     clicks=sum(clicks,na.rm=TRUE),
                     impressions=sum(impressions,na.rm=TRUE),
                     engagement=sum(engagement,na.rm=TRUE),
                     pr_value_rmb=sum(pr_value_rmb,na.rm=TRUE),
                     views=sum(views,na.rm=TRUE))]

#--------------define functions------------------------#
# file_name<-decomp_file[3]
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
  key_names<-c("var_ugg.cn", "var_tmall", "var_whs", "var_dtc")
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


ugg.cn<-sales_decompose(decomp_file[grepl(pattern = "UGG.CN",decomp_file)])
ugg.cn<-re_distribute(my_raw_decomp=ugg.cn)
# ugg.cn_spd<-map_spd(my_dataframe = ugg.cn, spd_data = spending)

tmall<-sales_decompose(decomp_file[grepl(pattern = "Tmall",decomp_file)])
tmall<-re_distribute(my_raw_decomp=tmall)
# tmall_spd<-map_spd(my_dataframe = tmall, spd_data = spending)


whs<-sales_decompose(decomp_file[grepl(pattern = "WHS",decomp_file)])
whs<-re_distribute(my_raw_decomp=whs)
# whs_spd<-map_spd(my_dataframe = whs, spd_data = spending)

dtc<-sales_decompose(decomp_file[grepl(pattern = "DTC",decomp_file)])
dtc<-re_distribute(my_raw_decomp=dtc)
# dtc_spd<-map_spd(my_dataframe = dtc, spd_data = spending)

ugg_decomp<-bind_rows(ugg.cn,tmall,whs, dtc)
ugg_decomp<-ugg_decomp[, metric:= "Sales Unit"]
ugg_decomp<-ugg_decomp[grepl("Traffic",Prodname),metric:= "Traffic"]
# ugg_decomp<-ugg_decomp[,channel_group:="Retail"]
# ugg_decomp<-ugg_decomp[channel=="UGG.CN"|channel=="Tmall",channel_group:="Online"]
ugg_decomp_online<-ugg_decomp[channel=="UGG.CN"|channel=="Tmall",by=.(week,
                                                                      Prodlvl, Geoglvl,
                                                                      variable,season, IPweight,
                                                                      Prodname,var_map, group2, group3,
                                                                      group4, var_key, year, mat, metric),
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
                                bumpup_factor=mean(bumpup_factor, na.rm=TRUE)
                                )]

ugg_decomp_online<-ugg_decomp_online[, c("channel","Geogname"):=list("Online", "Overall")]

ugg_decomp_retail<-ugg_decomp[channel!="UGG.CN"&channel!="Tmall",by=.(week,
                                                                      Prodlvl, Geoglvl,
                                                                      variable,season, IPweight,
                                                                      Prodname,var_map, group2, group3,
                                                                      group4, var_key, year, mat, metric),
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

ugg_decomp_retail<-ugg_decomp_retail[, c("channel","Geogname"):=list("Retail", "Overall")]
#-------------------------------------------------------------------------------------------#
ugg_decomp_whs<-ugg_decomp[channel=="WHS",by=.(week,
                                               Prodlvl, Geoglvl,
                                               variable,season, IPweight,
                                               Prodname,var_map, group2, group3,
                                               group4, var_key, year, mat, metric),
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

ugg_decomp_whs<-ugg_decomp_whs[, c("channel","Geogname"):=list("WHS", "Overall")]

# head(ugg_decomp[!(channel%in% c("WHS", "UGG.CN","Tmall")),],2)
# head(ugg_decomp[(channel%in% c("UGG.CN")),],2)

dtc<-dtc[, metric:= fifelse(grepl("Traffic",Prodname),
                            "Traffic","Sales Unit")]

ugg_decomp_dtc<-dtc[, by=.(week,
                           Prodlvl, Geoglvl,
                           variable,season, IPweight,
                           Prodname,var_map, group2, group3,
                           group4, var_key, year, mat, metric),
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

ugg_decomp_dtc<-ugg_decomp_dtc[, c("channel","Geogname"):=list("DTC", "Overall")]




#----------------------------------------------------------------------------------------------#
ugg_decomp_all<-ugg_decomp[,by=.(week,
                                 Prodlvl, Geoglvl,
                                 variable,season, IPweight,
                                 var_map, group2, group3,
                                 group4, var_key, year, mat, metric),
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

ugg_decomp_all<-ugg_decomp_all[, c("channel","Geogname"):=list("Overall", "Overall")]

ugg_decomp<-bind_rows(ugg_decomp_all, ugg_decomp_retail, ugg_decomp_online, 
                      ugg_decomp_dtc, ugg_decomp_whs,
                      ugg_decomp)



###aggregate
ugg_decomp_agg<-ugg_decomp[,by=.(metric, mat, season, 
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


# ugg_decomp_mapped_spd<-bind_rows(ugg.cn_spd, tmall_spd, whs_spd, dtc_spd)
# ugg_decomp_mapped_spd<-ugg_decomp_mapped_spd[, metric:= fifelse(grepl("Traffic",Prodname),
#                                           "Traffic","Sales Unit")]




# ugg_decomp_mapped_spd<-ugg_decomp_mapped_spd[,channel_group:="Retail"]
# ugg_decomp_mapped_spd<-ugg_decomp_mapped_spd[channel=="UGG.CN"|channel=="Tmall",
#                                              channel_group:="Online"]

#------------------------map spending--------------------------#

ugg_decomp_fit<-ugg_decomp[variable=="_Actual"|variable=="_FinlFit", 
                           by=.(channel, Geogname,product, 
                                week, season, year, mat, var_map, metric),
                           .(unitbump=sum(decomp_unit_bump, na.rm=TRUE),
                             netvalue=sum(decomp_NetRevenue,na.rm=TRUE),
                            grossvalue=sum(decomp_value,na.rm=TRUE),
                            unit=sum(decomp_unit, na.rm=TRUE))]%>%
  pivot_wider(names_from = var_map, values_from = c(unitbump, netvalue,grossvalue, unit))%>%
  arrange(channel,metric , product, Geogname, week)

ugg_decomp_fit<-setDT(ugg_decomp_fit)

#-------------------chart, predict vs. actual-----------------#
# online_chart<-ugg_decomp_fit%>%
#   filter(channel=="UGG.CN"|channel=="Tmall")%>%
#   ggplot(aes(x=week))+
#   geom_line(aes(y=unit_Actual),colour="darkred", size=1)+
#   geom_line(aes(y=unit_Predict), colour="steelblue", linetype="twodash", size=1)+
#   facet_wrap(channel~product, scales="free_y")
# 
# 
# whs_chart_sales<-ugg_decomp_fit%>%
#   filter(channel=="WHS"& metric=="Sales Unit" )%>%
#   ggplot(aes(x=week))+
#   geom_line(aes(y=unit_Actual),colour="darkred", size=1)+
#   geom_line(aes(y=unit_Predict), colour="steelblue", linetype="twodash", size=1)+
#   facet_wrap(channel~product, scales="free_y", ncol = 2)
# 
# whs_chart_traffic<-ugg_decomp_fit%>%
#   filter(channel=="WHS"& metric=="Traffic" )%>%
#   ggplot(aes(x=week))+
#   geom_line(aes(y=unit_Actual),colour="darkred", size=1)+
#   geom_line(aes(y=unit_Predict), colour="steelblue", linetype="twodash", size=1)+
#   facet_wrap(channel~product, scales="free_y",ncol = 2)
# 
# dtc_chart_sales<-ugg_decomp_fit%>%
#   filter(channel=="DTC"& metric=="Sales Unit")%>%
#   ggplot(aes(x=week))+
#   geom_line(aes(y=unit_Actual),colour="darkred", size=1)+
#   geom_line(aes(y=unit_Predict), colour="steelblue", linetype="twodash", size=1)+
#   facet_wrap(channel~product, scales="free_y",ncol = 2)
# 
# dtc_chart_traffic<-ugg_decomp_fit%>%
#   filter(channel=="DTC"& metric=="Traffic")%>%
#   ggplot(aes(x=week))+
#   geom_line(aes(y=unit_Actual),colour="darkred", size=1)+
#   geom_line(aes(y=unit_Predict), colour="steelblue", linetype="twodash", size=1)+
#   facet_wrap(channel~product, scales="free_y",ncol = 2)
# 
# pdf(paste(my_dir,"results","predict_vs_actual.pdf", sep="/"))
# print(online_chart)
# print(whs_chart_sales)
# print(whs_chart_traffic)
# print(dtc_chart_sales)
# print(dtc_chart_traffic)
# dev.off()

#-------------------write xlsx， first time --------------------------------#
# results<-createWorkbook()
# addWorksheet(results, sheetName = "baseline_decomp")
# addWorksheet(results, sheetName = "baseline_decomp_spd")
# addWorksheet(results, sheetName = "spending")
# addWorksheet(results, sheetName = "actual_sales")
# 
# writeDataTable(results, "baseline_decomp",x=ugg_decomp)
# writeDataTable(results, "baseline_decomp_spd",x=ugg_decomp_mapped_spd)
# writeDataTable(results, "spending",x=spending)
# writeDataTable(results, "actual_sales",x=ugg_decomp_fit)
# 
# saveWorkbook(results, paste(my_dir,"results", "ugg_baseline_decomp_test.xlsx", sep="/"),
#              overwrite = TRUE)

setcolorder(ugg_decomp, neworder = c("week","Prodlvl","Geoglvl","variable","season",
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

# head(ugg_decomp)

# Sys.setlocale(locale = "chs")
celebrity<-read_excel(paste(my_dir,"time map.xlsx", sep="/"),
                     sheet = "celebrity")%>%
  mutate(week=ymd(week))

celebrity<-setDT(celebrity)
celebrity_results<-merge(ugg_decomp, celebrity[,-2], by.x= "week", by.y="week")




# setcolorder(ugg_decomp_agg, neworder = c("metric","mat" ,"season" ,"channel",
#                                          "Geogname","Prodlvl" ,"Geoglvl" , "variable",
#                                      "IPweight", "var_map" , "group2" ,"group3",
#                                      "group4", "var_key", "year", "Prodname",
#                                      "geogkey", "prodkey" ,"BRAND","product",
#                                      "decomp_unit", "net_price", "decomp_NetRevenue", 
#                                      "decomp_unit_adjust",
#                                      "decomp_NetRevenue_adjust","traffic_unit",
#                                      "traffic_mkt",
#                                      "sold_price" ,"decomp_value" ,"decomp_value_adjust"))



print("!calc generating finished")



#-------------------write csv， the other time --------------------------------#
write_csv(ugg_decomp, paste(my_dir,"results", "ugg_baseline_decomp.csv", sep="/"))
# write_csv(ugg_decomp_mapped_spd, paste(my_dir,"results", "ugg_baseline_decomp_with_spd.csv", sep="/"))
write_csv(ugg_decomp_fit, paste(my_dir,"results", "ugg_actual_predict.csv", sep="/"))
write_csv(ugg_decomp_agg, paste(my_dir,"results", "ugg_actual_agg.csv", sep="/"))
# write_csv(celebrity_results, paste(my_dir,"results", "celebrity.csv", sep="/"))


print("!calc generating finished")

celebrity_excel<-createWorkbook()
addWorksheet(celebrity_excel, sheetName = "celebrity")

writeDataTable(celebrity_excel, "celebrity",x=celebrity_results)

saveWorkbook(celebrity_excel, paste(my_dir,"results", "celebrity.xlsx", sep="/"),
             overwrite = TRUE)


# write_csv(spending, paste(my_dir,"results", "spending.csv", sep="/"))

ym_decomp<-list.files(paste(my_dir, "decomp","Tmall ym" ,sep="/"), 
                      full.names = TRUE, pattern = "*xlsx")

tmall_ym<-sales_decompose(ym_decomp)
tmall_ym<-re_distribute(my_raw_decomp=tmall_ym)
# tmall_spd_ym<-map_spd(my_dataframe = tmall_ym, spd_data = spending)

tmall_ym<-tmall_ym[, metric:= "Sales Unit"]
tmall_ym<-tmall_ym[grepl("Traffic",Prodname),metric:= "Traffic"]

setcolorder(tmall_ym, neworder = c("week","Prodlvl","Geoglvl","variable","season",
                                   "IPweight", "var_map" , "group2" ,"group3",
                                   "group4", "var_key", "year", "mat", "metric",
                                   "decomp_unit", "sold_price", "decomp_value", 
                                   "decomp_unit_adjust",
                                   "decomp_value_adjust","traffic_unit",
                                   "traffic_mkt","channel", "Geogname", "Prodname",
                                   "geogkey","prodkey", "BRAND", "product",
                                   "traffic_contri","actual_weekly","traffic_weekly"))



#---------------------------Write Tmall----------------#
write_csv(tmall_ym, paste(my_dir,"results", "tmall_baseline_decomp.csv", sep="/"))
# write_csv(tmall_spd, paste(my_dir,"results", "tmall_baseline_decomp_with_spd.csv", sep="/"))
# write_csv(redistributed_decomp, paste(my_dir,"results", "tmall_baseline_decomp.csv", sep="/"))

print("!calc generating of Tmall finished")

write_csv(ugg.cn, paste(my_dir,"results", "ugg.cn_baseline_decomp.csv", sep="/"))


#--------------------------------------#
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
  sales_traffic_mkt<-sales_decomp_traffic[!(group4 %in%
                                              c("Base","Actual", "Predict", "Residual")),
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
                      sales_decomp_traffic[!(group4 %in%
                                               c("Base","Actual", "Predict", "Residual")),..my_cols],
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


#############
whs_covid<-read_excel("C:/005 Decker/005 Decker/Raw Data/WHS_Retailer_Partner/疫情期间销售追踪_0317 - Sarah - 01122021.xlsx",
                      sheet = "master", skip=3)


names(whs_covid)[1:6]<-c("partner", "store_id", "store_name_en",
                         "city", "open_date", "close_date")

whs_covid<-whs_covid%>%
  filter(!is.na(store_id))

# unique(whs_covid$store_name_en)
# names(whs_covid)<-as.character(names(whs_covid))
# names<-paste("week", 1:69, sep = "_")
# names<-data.frame(map_week=names, week=names(whs_covid)[7:75])

# names(whs_covid)[7:75]<-names$map_week
whs_covid<-whs_covid%>%
  mutate_at(.vars = 7:75, ~as.character(.))


whs_covid_long<-setDT(whs_covid)%>%
  pivot_longer(cols = 7:75, names_to = "map_week")

head(whs_covid_long)

whs_covid_0113<-createWorkbook()
addWorksheet(whs_covid_0113, sheetName = "whs")

writeDataTable(whs_covid_0113, "whs",x=whs_covid_long)

saveWorkbook(whs_covid_0113, "C:/005 Decker/005 Decker/Raw Data/WHS_Retailer_Partner/whs_covid_0113.xlsx",
             overwrite = TRUE)







