---
title: Axis.BaseUnit property (Excel)
keywords: vbaxl10.chm561104
f1_keywords:
- vbaxl10.chm561104
ms.prod: excel
api_name:
- Excel.Axis.BaseUnit
ms.assetid: f6fead0e-fc3f-834c-9a80-ae836b4f97d1
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.BaseUnit property (Excel)

Returns or sets the base unit for the specified category axis. Read/write **[XlTimeUnit](Excel.XlTimeUnit.md)**.


## Syntax

_expression_.**BaseUnit**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

Setting this property has no visible effect if the **[CategoryType](Excel.Axis.CategoryType.md)** property for the specified axis is set to **xlCategoryScale**. The set value is retained, however, and takes effect when the **CategoryType** property is set to **xlTimeScale**.

You cannot set this property for a value axis.


## Example

This example sets the category axis in embedded chart one on worksheet one to use a time scale, with months as the base unit.

```vb
With Worksheets(1).ChartObjects(1).Chart 
 With .Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnit = xlMonths 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]