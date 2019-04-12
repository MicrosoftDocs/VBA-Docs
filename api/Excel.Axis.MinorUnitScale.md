---
title: Axis.MinorUnitScale property (Excel)
keywords: vbaxl10.chm561107
f1_keywords:
- vbaxl10.chm561107
ms.prod: excel
api_name:
- Excel.Axis.MinorUnitScale
ms.assetid: bcbb3e11-5a30-f275-1beb-8575bac3a7fb
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.MinorUnitScale property (Excel)

Returns or sets the minor unit scale value for the category axis when the **CategoryType** property is set to **xlTimeScale**. Read/write **XlTimeUnit**.


## Syntax

_expression_.**MinorUnitScale**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

**XlTimeUnit** can be one of the **[XlTimeUnit](Excel.XlTimeUnit.md)** constants.


## Example

This example sets the category axis to use a time scale, and sets the major and minor units.

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