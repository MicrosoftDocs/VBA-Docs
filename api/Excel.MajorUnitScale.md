---
title: MajorUnitScale property (Excel Graph)
keywords: vbagr10.chm67185
f1_keywords:
- vbagr10.chm67185
ms.prod: excel
api_name:
- Excel.MajorUnitScale
ms.assetid: b2a54ca7-6eac-5552-6de7-ee0ab59e1ddb
ms.date: 04/11/2019
localization_priority: Normal
---


# MajorUnitScale property (Excel Graph)

Returns or sets the major unit scale value for the category axis when the **[CategoryType](excel.categorytype.md)** property is set to **xlTimeScale** (**[XlCategoryType](excel.xlcategorytype.md)**). Read/write **[XlTimeUnit](excel.xltimeunit.md)**.

## Syntax

_expression_.**MajorUnitScale**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the category axis to use a time scale, and sets the major and minor units.

```vb
With myChart.Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .MajorUnit = 5 
 .MajorUnitScale = xlDays 
 .MinorUnit = 1 
 .MinorUnitScale = xlDays 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]