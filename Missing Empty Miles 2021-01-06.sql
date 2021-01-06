DECLARE @DateStart DATETIME, 
@DateEnd DATETIME, 
@Branch VARCHAR(12)

SET @DateStart = '12-27-2020'  --DATEADD(dd, -(datepart(dw, GetDate())+7), GetDate())
SET @DateEnd = '01-02-2021' --DATEADD(dd, -datepart(dw, GetDate()+1), GetDate())
SET @Branch = '580' 

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
[Last Update]=FORMAT(lgh.lgh_updatedon,'MM/dd/yyyy HH:mm'),
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
AND trc.trc_branch = @Branch
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
[Last Update]=FORMAT(lgh.lgh_updatedon,'MM/dd/yyyy HH:mm'),
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
AND trc.trc_branch = @Branch
AND lgh.lgh_outstatus = 'CMP'
)
AS a
WHERE a.MissingOrder <> 0)