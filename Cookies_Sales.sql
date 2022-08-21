

-- shows all orders details for year 2020
SELECT 
b.[Name],a.Product,a.Units_Sold,a.[Date]
FROM Orders as a left join Customers as b on a.Customer_ID=b.Customer_ID
where Year([Date])='2020'
order by a.[Date]

-- shows all orders revenue for year 2020
SELECT 
a.Product,a.Units_Sold,(a.Units_Sold * b.Cost_Per_Cookie) as Revenue,a.[Date]
FROM Orders as a join CookieTypes as b on a.Customer_ID=b.Cookie_Type
where Year([Date])='2020'
order by a.[Date]



-- Shows product total units sold per month for year 2020 
SELECT 
a.Product,sum(a.Units_Sold) as Units_Sold,a.[Date]
FROM Orders as a left join Customers as b on a.Customer_ID=b.Customer_ID
where Year([Date])='2020'
group by a.Product,MONTH(a.[Date]),a.[Date]
order by a.[Date]


