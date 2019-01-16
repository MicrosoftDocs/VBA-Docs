---
title: ODSOFilter.CompareTo property (Office)
keywords: vbaof11.chm240005
f1_keywords:
- vbaof11.chm240005
ms.prod: office
api_name:
- Office.ODSOFilter.CompareTo
ms.assetid: dc14c506-1315-d0f9-edcd-38c395feab63
ms.date: 06/08/2017
localization_priority: Normal
---


# ODSOFilter.CompareTo property (Office)

Gets or sets a  **String** that represents the text to compare in the query filter criterion. Read/write.


## Syntax

_expression_. `CompareTo`

_expression_ A variable that represents an [ODSOFilter](Office.ODSOFilter.md) object.


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


## See also


[ODSOFilter Object](Office.ODSOFilter.md)



[ODSOFilter Object Members](./overview/Library-Reference/odsofilter-members-office.md)

