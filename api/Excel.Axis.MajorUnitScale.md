---
title: Axis.MajorUnitScale property (Excel)
keywords: vbaxl10.chm561106
f1_keywords:
- vbaxl10.chm561106
ms.prod: excel
api_name:
- Excel.Axis.MajorUnitScale
ms.assetid: f0f4b179-f166-4fe6-f333-365edc5bc4f7
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.MajorUnitScale property (Excel)

Returns or sets the major unit scale value for the category axis when the **CategoryType** property is set to **xlTimeScale**. Read/write **XlTimeUnit**.


## Syntax

_expression_.**MajorUnitScale**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

**XlTimeUnit** can be one of the **[XlTimeUnit](Excel.XlTimeUnit.md)** constants.

## Example

This example sets the category axis to use a time scale and sets the major and minor units.

```vb
With Charts(1).Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .MajorUnit = 5 
 .MajorUnitScale = xlDays 
 .MinorUnit = 1 
 .MinorUnitScale = xlDays 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
