---
title: BaseUnit property (Excel Graph)
keywords: vbagr10.chm3076962
f1_keywords:
- vbagr10.chm3076962
ms.prod: excel
api_name:
- Excel.BaseUnit
ms.assetid: 05c83ae8-ab67-1330-3a78-f0219e72637a
ms.date: 04/09/2019
localization_priority: Normal
---


# BaseUnit property (Excel Graph)

Returns or sets the base unit for the specified category axis. Read/write **[XlTimeUnit](excel.xltimeunit.md)**.

## Syntax

_expression_.**BaseUnit**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

Setting this property has no visible effect if the **[CategoryType](excel.categorytype.md)** property for the specified axis is set to **xlCategoryScale**. The set value is retained, however, and takes effect when the **CategoryType** property is set to **xlTimeScale**.

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