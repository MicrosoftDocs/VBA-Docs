---
title: ODSOFilter.Column property (Office)
keywords: vbaof11.chm240003
f1_keywords:
- vbaof11.chm240003
ms.prod: office
api_name:
- Office.ODSOFilter.Column
ms.assetid: 53caf4f7-73f1-3969-b407-8fa89883c78d
ms.date: 06/08/2017
localization_priority: Normal
---


# ODSOFilter.Column property (Office)

Gets or sets a  **String** that represents the name of the field in the mail merge data source to use in the filter. Read/write.


## Syntax

_expression_. `Column`

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
[TextFrame2 Object](Office.TextFrame2.md)



[ODSOFilter Object Members](./overview/Library-Reference/odsofilter-members-office.md)
[TextFrame2 Object Members](./overview/Library-Reference/textframe2-members-office.md)

