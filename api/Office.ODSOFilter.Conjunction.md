---
title: ODSOFilter.Conjunction property (Office)
keywords: vbaof11.chm240006
f1_keywords:
- vbaof11.chm240006
ms.prod: office
api_name:
- Office.ODSOFilter.Conjunction
ms.assetid: 22d2287c-9b0e-c4ce-164d-e8424c62aa86
ms.date: 06/08/2017
localization_priority: Normal
---


# ODSOFilter.Conjunction property (Office)

Gets or sets an  **MsoFilterConjunction** constant that represents how a filter criterion relates to other filter criteria in the **ODSOFilters** object. Read/write.


## Syntax

_expression_. `Conjunction`

_expression_ A variable that represents an [ODSOFilter](Office.ODSOFilter.md) object.


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a  **Region** field equal to "WA".


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]