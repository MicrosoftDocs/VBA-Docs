---
title: ODSOFilter.Comparison property (Office)
keywords: vbaof11.chm240004
f1_keywords:
- vbaof11.chm240004
ms.prod: office
api_name:
- Office.ODSOFilter.Comparison
ms.assetid: 992565b3-90c5-4f44-7cae-ba0533529127
ms.date: 01/22/2019
localization_priority: Normal
---


# ODSOFilter.Comparison property (Office)

Gets or sets an **[MsoFilterComparison](office.msofiltercomparison.md)** constant that represents how to compare the **Column** and **CompareTo** properties. Read/write.


## Syntax

_expression_.**Comparison**

_expression_ A variable that represents an **[ODSOFilter](Office.ODSOFilter.md)** object.


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a **Region** field equal to "WA".


```vb
Sub SetQueryCriterion() 
 Dim appOffice As Office.OfficeDataSourceObject 
 Dim intItem As Integer 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _ 
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

- [ODSOFilter object members](overview/library-reference/odsofilter-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

