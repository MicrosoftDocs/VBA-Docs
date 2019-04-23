---
title: CategoryType property (Excel Graph)
keywords: vbagr10.chm67187
f1_keywords:
- vbagr10.chm67187
ms.prod: excel
api_name:
- Excel.CategoryType
ms.assetid: 6af3b261-abed-a78a-5952-645af07cde9d
ms.date: 04/10/2019
localization_priority: Normal
---


# CategoryType property (Excel Graph)

Returns or sets the category axis type. Read/write **[XlCategoryType](excel.xlcategorytype.md)**.

## Syntax

_expression_.**CategoryType**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

You cannot set this property for a value axis.


## Example

This example sets the category axis on the chart to use a time scale, with months as the base unit.

```vb
With myChart 
 With .Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnit = xlMonths 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]