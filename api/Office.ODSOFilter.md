---
title: ODSOFilter object (Office)
keywords: vbaof11.chm240000
f1_keywords:
- vbaof11.chm240000
ms.prod: office
api_name:
- Office.ODSOFilter
ms.assetid: 9c1babb7-31af-3c43-47ae-3864f6462c27
ms.date: 06/08/2017
localization_priority: Normal
---


# ODSOFilter object (Office)

Represents a filter to be applied to an attached mail merge data source. The  **ODSOFilter** object is a member of the **ODSOFilters** object.


## Remarks

Each filter is a line in a query string. Use the  **Column**, **Comparison**, **CompareTo**, and **Conjunction** properties to return or set the data source query criterion.


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA".


```vb
Sub SetQueryCriterion() 
 Dim appOffice As Office.OfficeDataSourceObject 
 Dim intItem As Integer 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Filters 
 For intItem = 1 To .Count 
 With .Item(intItem) 
 If .Column = "Region" Then 
 .Comparison = msoFilterComparisonNotEqual 
 .CompareTo = "WA" 
 If .Conjunction = "Or" Then .Conjunction = "And" 
 End If 
 End With 
 Next intItem 
 End With 
End Sub
```

Use the  **Add** method of the **ODSOFilters** object to add a new filter criterion to the query. This example adds a new line to the query string and then applies the combined filter to the data source.




```vb
Sub SetQueryCriterion() 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Filters 
 .Add Column:="Region", _ 
 Comparison:=msoFilterComparisonIsBlank, _ 
 Conjunction:=msoFilterConjunctionAnd 
 .ApplyFilter 
 End With 
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Office.ODSOFilter.Application.md)|
|[Column](Office.ODSOFilter.Column.md)|
|[CompareTo](Office.ODSOFilter.CompareTo.md)|
|[Comparison](Office.ODSOFilter.Comparison.md)|
|[Conjunction](Office.ODSOFilter.Conjunction.md)|
|[Creator](Office.ODSOFilter.Creator.md)|
|[Index](Office.ODSOFilter.Index.md)|
|[Parent](Office.ODSOFilter.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]