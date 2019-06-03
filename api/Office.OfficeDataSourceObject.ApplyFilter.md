---
title: OfficeDataSourceObject.ApplyFilter method (Office)
keywords: vbaof11.chm232009
f1_keywords:
- vbaof11.chm232009
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.ApplyFilter
ms.assetid: 9ce3ed9b-3d84-9ebd-68ae-be958ad1a99c
ms.date: 01/22/2019
localization_priority: Normal
---


# OfficeDataSourceObject.ApplyFilter method (Office)

Applies a filter to a mail merge data source to filter specified records meeting specified criteria.


## Syntax

_expression_.**ApplyFilter**

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