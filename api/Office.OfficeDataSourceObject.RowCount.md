---
title: OfficeDataSourceObject.RowCount property (Office)
keywords: vbaof11.chm232005
f1_keywords:
- vbaof11.chm232005
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.RowCount
ms.assetid: 5360a399-e2f8-b331-f62c-c110884b3c92
ms.date: 01/22/2019
localization_priority: Normal
---


# OfficeDataSourceObject.RowCount property (Office)

Gets a **Long** that represents the number of records in the specified data source. Read-only.


## Syntax

_expression_.**RowCount**

_expression_ A variable that represents an **[OfficeDataSourceObject](Office.OfficeDataSourceObject.md)** object.


## Example

This example adds a new filter that removes all records with a blank **Region** field and then applies the filter to the active publication.


```vb
Sub OfficeFilters() 
 Dim appOffice As OfficeDataSourceObject 
 Dim appFilters As ODSOFilters 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 Set appFilters = appOffice.Filters 
 
 MsgBox appOffice.RowCount 
 
 appFilters.Add Column:="Region", Comparison:=msoFilterComparisonEqual, _ 
 Conjunction:=msoFilterConjunctionAnd, bstrCompareTo:="WA" 
 appOffice.ApplyFilter 
 
 MsgBox appOffice.RowCount 
 
End Sub
```


## See also

- [OfficeDataSourceObject object members](overview/library-reference/officedatasourceobject-members-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]