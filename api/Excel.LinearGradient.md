---
title: LinearGradient object (Excel)
keywords: vbaxl10.chm854072
f1_keywords:
- vbaxl10.chm854072
ms.prod: excel
api_name:
- Excel.LinearGradient
ms.assetid: cb648564-0f57-f1b9-1c89-0329c110583f
ms.date: 03/30/2019
localization_priority: Normal
---


# LinearGradient object (Excel)

The **LinearGradient** object transitions through a series of colors in a linear manner along a specific angle.


## Remarks

Attempting to access the **[Gradient](excel.interior.gradient.md)** property of an **Interior** object that does not have an existing gradient fill results in a run-time error. Be aware of the **[Pattern](Excel.Interior.Pattern.md)** property before accessing the **Gradient** property.
    
If the **Pattern** property is changed from a gradient type to a non-gradient type, the **Gradient** property will populate with default values.


## Properties

- [Application](Excel.LinearGradient.Application.md)
- [ColorStops](Excel.LinearGradient.ColorStops.md)
- [Creator](Excel.LinearGradient.Creator.md)
- [Degree](Excel.LinearGradient.Degree.md)
- [Parent](Excel.LinearGradient.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]