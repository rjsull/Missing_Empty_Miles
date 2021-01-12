#Import Libraries
library(tidyverse)
library(RODBC)
library(openxlsx)

###Connect to TMW Suite Replication
dbhandle <- odbcDriverConnect('driver={SQL Server};server=NFIV-SQLTMW-04;database=TMWSuite;trusted_connection=true')

###Contains SQL for finding what moves end and start in different cities. Ops uses this to add empty mile segements to orders that way the drivers get paid for the miles driven. Goal is to have no missing empty miles.
missingEmptyMiles <- sqlQuery(dbhandle, "
DECLARE @DateStart DATETIME, 
@DateEnd DATETIME, 
@Branch1 VARCHAR(12),
@Branch2 VARCHAR(12),
@Branch3 VARCHAR(12),
@Branch4 VARCHAR(12),
@Branch5 VARCHAR(12),
@Branch6 VARCHAR(12),
@Branch7 VARCHAR(12)


SET @DateStart = 
	CASE 
	WHEN DATEPART(DW,GETDATE()) = 4 OR DATEPART(DW,GETDATE()) = 5 OR DATEPART(DW,GETDATE()) = 6 OR DATEPART(DW,GETDATE()) = 7 --On Wed, Thu, and Fri Mornings, send Start of Current Week
	THEN DATEADD(wk, 0, DATEADD(DAY, 1-DATEPART(WEEKDAY, GETDATE()), DATEDIFF(dd, 0, GETDATE()))) --first day current week
	ELSE DATEADD(wk, -1, DATEADD(DAY, 1-DATEPART(WEEKDAY, GETDATE()), DATEDIFF(dd, 0, GETDATE()))) --first day previous week
	END
SET @DateEnd = 	
	CASE 
	WHEN DATEPART(DW,GETDATE()) = 4 OR DATEPART(DW,GETDATE()) = 5 OR DATEPART(DW,GETDATE()) = 6  OR DATEPART(DW,GETDATE()) = 7 --On Wed, Thu, and Fri Mornings, send End of Current Week
	THEN DATEADD(wk, 1, DATEADD(DAY, 0-DATEPART(WEEKDAY, GETDATE()), DATEDIFF(dd, 0, GETDATE()))) --last day current week
	ELSE DATEADD(wk, 0, DATEADD(DAY, 0-DATEPART(WEEKDAY, GETDATE()), DATEDIFF(dd, 0, GETDATE()))) --last day previous week
	END
SET @Branch1 = '570'
SET @Branch2 = '571'
SET @Branch3 = '572'
SET @Branch4 = '573'
SET @Branch5 = '574'
SET @Branch6 = '580'
SET @Branch7 = '581'

SELECT a.*
FROM (
SELECT [Tractor Branch]=trc.trc_branch, 
Tractor = lgh.lgh_tractor, 
[Driver Name]=lgh.lgh_driver1, 
[Trailer] = lgh.lgh_primary_trailer, 
[Bill To] = ord.ord_billto, 
[OrderNumber] = ord.ord_hdrnumber, 
[Leg Number] = lgh.lgh_number, 
[Move Number]=lgh.mov_number, 
[Stop Count]=ISNULL(ord.ord_stopcount,0), 
[Start Date]=FORMAT(lgh.lgh_startdate,'MM/dd/yyyy HH:mm'), 
[Start Location]=lgh.cmp_id_start, 
[Start City State]=ctyStart.cty_name + ', ' + ctyStart.cty_state, 
[P/U Date]=FORMAT(lgh.lgh_rstartdate,'MM/dd/yyyy HH:mm'), 
Shipper=lgh.cmp_id_rstart, 
[Shipper City State]=ctyShip.cty_name + ', ' + ctyShip.cty_state, 
[Delivery Date]=FORMAT(lgh.lgh_renddate,'MM/dd/yyyy HH:mm'), 
Consignee=cmp_id_rend, 
[Delivery City State]=ctyCon.cty_name + ', ' + ctyCon.cty_state, 
[End Date]=FORMAT(lgh.lgh_enddate,'MM/dd/yyyy HH:mm'), 
[End Location]=lgh.cmp_id_end, 
[End City State]=ctyEnd.cty_name + ', ' + ctyEnd.cty_state, 
[Leg Miles] = lgh.lgh_miles, 
[Ref#]=ord.ord_refnum, 
[Update User]=lgh.lgh_updatedby, 
[Last Update]=lgh.lgh_updatedon,
[Fleet Manager] = mpp.mpp_teamleader,
[MissingOrder] = 
	CASE 
	WHEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) <> ctyEnd.cty_name + ', ' + ctyEnd.cty_state
		AND LEAD(lgh.lgh_driver1,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) = lgh.lgh_driver1
	THEN LEAD(ord.ord_hdrnumber,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate)
	ELSE 0
	END,
[MissingStart] = 
	CASE 
	WHEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) <> ctyEnd.cty_name + ', ' + ctyEnd.cty_state
		AND LEAD(lgh.lgh_driver1,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) = lgh.lgh_driver1
	THEN ctyEnd.cty_name + ', ' + ctyEnd.cty_state
	ELSE ''
	END,
[MissingEnd] = 
	CASE 
	WHEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) <> ctyEnd.cty_name + ', ' + ctyEnd.cty_state
		AND LEAD(lgh.lgh_driver1,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) = lgh.lgh_driver1
	THEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate)
	ELSE ''
	END
FROM legheader lgh LEFT JOIN orderheader ord ON lgh.ord_hdrnumber = ord.ord_hdrnumber 
LEFT JOIN tractorprofile trc ON lgh.lgh_tractor = trc.trc_number 
LEFT JOIN city ctyStart ON lgh.lgh_startcity = ctyStart.cty_code 
LEFT JOIN city ctyEnd ON lgh.lgh_endcity = ctyEnd.cty_code 
LEFT JOIN city ctyShip ON ord.ord_origincity = ctyShip.cty_code 
LEFT JOIN city ctyCon ON ord.ord_destcity = ctyCon.cty_code 
LEFT JOIN manpowerprofile mpp ON lgh.lgh_driver1 = mpp.mpp_id
WHERE CONVERT(varchar(10),lgh.lgh_startdate,102) BETWEEN '' + CONVERT(varchar(10),@DateStart,102) + '' AND '' + CONVERT(varchar(10),@DateEnd,102) + ''
AND trc.trc_branch IN (@Branch1, @Branch2, @Branch3, @Branch4, @Branch5, @Branch6, @Branch7)
AND lgh.lgh_outstatus = 'CMP'
--AND lgh.lgh_driver1 IN ('CLEAN')
--ORDER BY lgh.lgh_driver1 ASC,
--lgh.lgh_startdate ASC
)
AS a
WHERE a.MissingOrder <> 0
OR OrderNumber IN (
SELECT DISTINCT a.MissingOrder
FROM (
SELECT [Tractor Branch]=trc.trc_branch, 
Tractor = lgh.lgh_tractor, 
[Driver Name]=lgh.lgh_driver1, 
[Trailer] = lgh.lgh_primary_trailer, 
[Bill To] = ord.ord_billto, 
[OrderNumber] = ord.ord_hdrnumber, 
[Leg Number] = lgh.lgh_number, 
[Move Number]=lgh.mov_number, 
[Stop Count]=ISNULL(ord.ord_stopcount,0), 
[Start Date]=FORMAT(lgh.lgh_startdate,'MM/dd/yyyy HH:mm'), 
[Start Location]=lgh.cmp_id_start, 
[Start City State]=ctyStart.cty_name + ', ' + ctyStart.cty_state, 
[P/U Date]=FORMAT(lgh.lgh_rstartdate,'MM/dd/yyyy HH:mm'), 
Shipper=lgh.cmp_id_rstart, 
[Shipper City State]=ctyShip.cty_name + ', ' + ctyShip.cty_state, 
[Delivery Date]=FORMAT(lgh.lgh_renddate,'MM/dd/yyyy HH:mm'), 
Consignee=cmp_id_rend, 
[Delivery City State]=ctyCon.cty_name + ', ' + ctyCon.cty_state, 
[End Date]=FORMAT(lgh.lgh_enddate,'MM/dd/yyyy HH:mm'), 
[End Location]=lgh.cmp_id_end, 
[End City State]=ctyEnd.cty_name + ', ' + ctyEnd.cty_state, 
[Leg Miles] = lgh.lgh_miles, 
[Ref#]=ord.ord_refnum, 
[Update User]=lgh.lgh_updatedby, 
[Last Update]=lgh.lgh_updatedon,
[Fleet Manager] = mpp.mpp_teamleader,
[MissingOrder] = 
	CASE 
	WHEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) <> ctyEnd.cty_name + ', ' + ctyEnd.cty_state
		AND LEAD(lgh.lgh_driver1,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) = lgh.lgh_driver1
	THEN LEAD(ord.ord_hdrnumber,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate)
	ELSE 0
	END,
[MissingStart] = 
	CASE 
	WHEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) <> ctyEnd.cty_name + ', ' + ctyEnd.cty_state
		AND LEAD(lgh.lgh_driver1,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) = lgh.lgh_driver1
	THEN ctyEnd.cty_name + ', ' + ctyEnd.cty_state
	ELSE ''
	END,
[MissingEnd] = 
	CASE 
	WHEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) <> ctyEnd.cty_name + ', ' + ctyEnd.cty_state
		AND LEAD(lgh.lgh_driver1,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) = lgh.lgh_driver1
	THEN LEAD(ctyStart.cty_name,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate) + ', ' + LEAD(ctyStart.cty_state,1) OVER (ORDER BY lgh.lgh_driver1, lgh.lgh_startdate)
	ELSE ''
	END
FROM legheader lgh LEFT JOIN orderheader ord ON lgh.ord_hdrnumber = ord.ord_hdrnumber 
LEFT JOIN tractorprofile trc ON lgh.lgh_tractor = trc.trc_number 
LEFT JOIN city ctyStart ON lgh.lgh_startcity = ctyStart.cty_code 
LEFT JOIN city ctyEnd ON lgh.lgh_endcity = ctyEnd.cty_code 
LEFT JOIN city ctyShip ON ord.ord_origincity = ctyShip.cty_code 
LEFT JOIN city ctyCon ON ord.ord_destcity = ctyCon.cty_code 
LEFT JOIN manpowerprofile mpp ON lgh.lgh_driver1 = mpp.mpp_id
WHERE CONVERT(varchar(10),lgh.lgh_startdate,102) BETWEEN '' + CONVERT(varchar(10),@DateStart,102) + '' AND '' + CONVERT(varchar(10),@DateEnd,102) + ''
AND trc.trc_branch IN (@Branch1, @Branch2, @Branch3, @Branch4, @Branch5, @Branch6, @Branch7)
AND lgh.lgh_outstatus = 'CMP'
)
AS a
WHERE a.MissingOrder <> 0)
")

###Close DB connection
odbcClose(dbhandle)

###Create summary sheet
df <- data.frame(missingEmptyMiles)
df_filtered <- df %>% subset(MissingOrder!='0')
df_filtered <- df_filtered %>% select (1,26,2,3,6,27,10,28,29)
df_filtered<-df_filtered[order(df_filtered$Tractor.Branch, df_filtered$Driver.Name),]

###Create Excel File and add sheets with data  
wb <- createWorkbook(creator = ifelse(.Platform$OS.type == "windows", Sys.getenv("USERNAME"), Sys.getenv("USER")))
sheet1 <- "Missing Empty Miles Summary"
sheet2 <- "Leg Data"
n <- ncol(df_filtered)
n2 <- ncol(df)
addWorksheet(wb, sheet1)
addWorksheet(wb, sheet2)
writeData(wb, sheet1, df_filtered, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
writeData(wb, sheet2, df, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)


###Add PC Miler Formula with Header
Header <- c('TEXT("Miles",0)')
writeFormula(wb, sheet = 1, x = Header, startCol = 10, startRow = 1)
MilesFormula = paste(paste0("Miles("), paste(paste0("H", 1:nrow(df_filtered) + 1L), paste0("I", 1:nrow(df_filtered) + 1L), sep = ","), paste0(")"))
writeFormula(wb, sheet = 1, x = MilesFormula, startCol = 10, startRow = 2)


###Format sheets, add filters, freeze panes, auto fit columns 
addFilter(wb, sheet1, row = 1, cols = 1:n)
addFilter(wb, sheet2, row = 1, cols = 1:n2)
freezePane(wb, sheet1, firstRow = TRUE)
freezePane(wb, sheet2, firstRow = TRUE) 
setColWidths(wb, sheet1, cols = 1:n, widths = "auto")
setColWidths(wb, sheet2, cols = 1:n2, widths = "auto")

###Save workbook, change path to save file locally 
saveWorkbook(wb, file = "C:/Users/sullivanry/Documents/MissingEmptyMiles.xlsx", overwrite = TRUE)

