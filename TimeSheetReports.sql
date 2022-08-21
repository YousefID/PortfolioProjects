


-- Generate employee access report per region
declare @month int
DECLARE @year INT
SET @month =MONTH(DATEADD(m, -1, getdate()))
if (@month=12)
SET @year =YEAR(DATEADD(yyyy, -1, getdate())) 
else
SET @year =YEAR(DATEADD(yyyy, 0, getdate()))

Delete from Ext_M01
Insert into EXT_M01 
select a.emp_no,b.S_name +','+b.F_name,a.Time_Profile,b.sgn as Division, b.dept,b.Loc,
case when substring(b.Loc,1,3) ='JED' THEN 'Western' 
when substring(b.Loc,1,3) ='JED' THEN 'Western'
when substring(b.Loc,1,3) ='YNB' THEN 'Western'
when substring(b.Loc,1,3) ='RUH' THEN 'Central'
when substring(b.Loc,1,3) ='JUB' THEN 'Eastern'
when substring(b.Loc,1,3) ='DMA' THEN 'Eastern'
when substring(b.Loc,1,3) ='AKB' THEN 'Eastern'
 END AS REG,cast((cast(a.Total as decimal(18,2))/cast(a.Planned as decimal(18,2)))*100 as decimal(18,2)) as Availibility,
 CASE when BMON='1' THEN 'Jan' + ' '+ BYEAR
WHEN BMON=2 THEN 'Feb' + ' ' + BYEAR
WHEN BMON=3 THEN 'Mar'  + ' ' + BYEAR
WHEN BMON=4 THEN 'Apr' + ' ' + BYEAR 
WHEN BMON=5 THEN 'May' + ' ' + BYEAR
WHEN BMON=6 THEN 'Jun'  + ' ' + BYEAR
WHEN BMON=7 THEN 'Jul'  + ' ' + BYEAR
WHEN BMON=8 THEN 'Aug'  + ' ' + BYEAR
WHEN BMON=9 THEN 'Sep'  + ' ' + BYEAR
WHEN BMON=10 THEN 'Oct'  + ' ' + BYEAR
WHEN BMON=11 THEN 'Nov'  + ' ' + BYEAR
WHEN BMON=12 THEN 'Dec' + ' ' + BYEAR  end as MNTH,'SLSA'
From [TRA_Deduction] a  join [master_data].[dbo].[chcm] b on a.emp_no = b.emp_no
join add_emp_data c  on c.emp_no=a.emp_no
WHERE ( c.TA_X_Flag is null or c.TA_X_Flag<>'1') and  
b.ARE <>'4659'  and BMON=@month  and BYEAR=@year 





-- Generate employee access report per division

INSERT INTO [Ext_DivisionsReport]
SELECT     TOP (100) PERCENT Division, a_1.Location, a_1.MNTH,Category_1,Category_2,Category_3,Category_4,Category_5,Report_ID
FROM         (SELECT     a.DIV AS Division, a.REG AS Location,a.MNTH,CONVERT(float, COUNT(b.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_1,CONVERT(float, COUNT(c.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_2,CONVERT(float, COUNT(d.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_3,CONVERT(float, COUNT(e.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_4,CONVERT(float, COUNT(f.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_5,3 as [Report_ID]
					  FROM          dbo.Ext_M01 AS a LEFT OUTER JOIN
                                              dbo.Ext_M01 AS b ON a.ID = b.ID AND cast(b.availability as decimal(18,2)) > 100  
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS c ON a.ID = c.ID AND cast(c.availability as decimal(18,2)) > 90  and cast(c.availability as decimal(18,2)) <= 100
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS d ON a.ID = d.ID AND cast(d.availability as decimal(18,2)) > 78  and cast(d.availability as decimal(18,2)) <= 90
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS e ON a.ID = e.ID AND cast(e.availability as decimal(18,2)) > 67  and cast(e.availability as decimal(18,2)) <= 78
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS f ON a.ID = f.ID AND cast(f.availability as decimal(18,2)) < 67 
											
													  	
                                         
							 GROUP BY a.DIV, a.REG ,a.MNTH 
--Siemensany
UNION
SELECT     'All SLSA' AS Division, a.REG AS Location,a.MNTH,CONVERT(float, COUNT(b.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_1,CONVERT(float, COUNT(c.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_2,CONVERT(float, COUNT(d.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_3,CONVERT(float, COUNT(e.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_4,CONVERT(float, COUNT(f.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_5,3 as [Report_ID]
					   FROM          dbo.Ext_M01 AS a LEFT OUTER JOIN
                                              dbo.Ext_M01 AS b ON a.ID = b.ID AND cast(b.availability as decimal(18,2)) > 100  
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS c ON a.ID = c.ID AND cast(c.availability as decimal(18,2)) > 90  and cast(c.availability as decimal(18,2)) <= 100
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS d ON a.ID = d.ID AND cast(d.availability as decimal(18,2)) > 78  and cast(d.availability as decimal(18,2)) <= 90
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS e ON a.ID = e.ID AND cast(e.availability as decimal(18,2)) > 67  and cast(e.availability as decimal(18,2)) <= 78
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS f ON a.ID = f.ID AND cast(f.availability as decimal(18,2)) < 67 
											
													  	
                                         
							 GROUP BY  a.REG,a.MNTH       
UNION
SELECT     a.DIV AS Division, 'All' AS Location,a.MNTH,CONVERT(float, COUNT(b.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_1,CONVERT(float, COUNT(c.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_2,CONVERT(float, COUNT(d.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_3,CONVERT(float, COUNT(e.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_4,CONVERT(float, COUNT(f.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_5,3 as [Report_ID]
					   FROM          dbo.Ext_M01 AS a LEFT OUTER JOIN
                                              dbo.Ext_M01 AS b ON a.ID = b.ID AND cast(b.availability as decimal(18,2)) > 100  
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS c ON a.ID = c.ID AND cast(c.availability as decimal(18,2)) > 90  and cast(c.availability as decimal(18,2)) <= 100
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS d ON a.ID = d.ID AND cast(d.availability as decimal(18,2)) > 78  and cast(d.availability as decimal(18,2)) <= 90
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS e ON a.ID = e.ID AND cast(e.availability as decimal(18,2)) > 67  and cast(e.availability as decimal(18,2)) <= 78
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS f ON a.ID = f.ID AND cast(f.availability as decimal(18,2)) < 67 
											
													  	
                                         
							 GROUP BY  a.DIV,a.MNTH   

UNION
SELECT     'All SLSA' AS Division, 'All' AS Location,a.MNTH,CONVERT(float, COUNT(b.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_1,CONVERT(float, COUNT(c.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_2,CONVERT(float, COUNT(d.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_3,CONVERT(float, COUNT(e.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_4,CONVERT(float, COUNT(f.ID)) 
                                              / CONVERT(float, COUNT(a.ID)) * 100 AS Category_5,3 as [Report_ID]
					   FROM          dbo.Ext_M01 AS a LEFT OUTER JOIN
                                              dbo.Ext_M01 AS b ON a.ID = b.ID AND cast(b.availability as decimal(18,2)) > 100  
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS c ON a.ID = c.ID AND cast(c.availability as decimal(18,2)) > 90  and cast(c.availability as decimal(18,2)) <= 100
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS d ON a.ID = d.ID AND cast(d.availability as decimal(18,2)) > 78  and cast(d.availability as decimal(18,2)) <= 90
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS e ON a.ID = e.ID AND cast(e.availability as decimal(18,2)) > 67  and cast(e.availability as decimal(18,2)) <= 78
													  LEFT OUTER JOIN
                                              dbo.Ext_M01 AS f ON a.ID = f.ID AND cast(f.availability as decimal(18,2)) < 67 
											  GROUP BY  a.MNTH   ) as a_1

ORDER BY division
								

-- Shows duplicate records from Timesheet
select * from Ext_SheetEntry a
JOIN
(
     SELECT
            WORKDATE,
            EMPNO,
            ACTIVITY,
            WBS,CATSHOURS
          
     FROM
            Ext_SheetEntry
     GROUP BY WORKDATE, EMPNO, ACTIVITY, WBS,CATSHOURS
     HAVING COUNT(*) >1
) b
on a.WORKDATE =b.WORKDATE
and a.EMPNO = b.EMPNO
and a.WBS = b.WBS
and a.ACTIVITY = b.ACTIVITY
and a.CATSHOURS=b.CATSHOURS
													  	
                                         
							      
