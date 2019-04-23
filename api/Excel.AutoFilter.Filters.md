---
title: AutoFilter.Filters property (Excel)
keywords: vbaxl10.chm538074
f1_keywords:
- vbaxl10.chm538074
ms.prod: excel
api_name:
- Excel.AutoFilter.Filters
ms.assetid: 4a22dcab-4d06-01a8-7811-4590cf28f506
ms.date: 04/13/2019
localization_priority: Normal
---


# AutoFilter.Filters property (Excel)

Returns a **[Filters](Excel.Filters.md)** collection that represents all the filters in an autofiltered range. Read-only.


## Syntax

_expression_.**Filters**

_expression_ A variable that represents an **[AutoFilter](Excel.AutoFilter.md)** object.


## Example

The following example sets a variable to the value of the **Criteria1** property of the filter for the first column in the filtered range on the Crew worksheet.

```vb
With Worksheets("Crew") 
 If .AutoFilterMode Then 
 With .AutoFilter.Filters(1) 
 If .On Then c1 = .Criteria1 
 End With 
 End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
