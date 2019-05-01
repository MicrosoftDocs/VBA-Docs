---
title: RectangularGradient object (Excel)
keywords: vbaxl10.chm856072
f1_keywords:
- vbaxl10.chm856072
ms.prod: excel
api_name:
- Excel.RectangularGradient
ms.assetid: e668d158-0436-cb27-a6f5-e27453681d66
ms.date: 04/02/2019
localization_priority: Normal
---


# RectangularGradient object (Excel)

The **RectangularGradient** object transitions through a series of colors in a linear manner along a specific angle.


## Remarks

Attempting to access a **[Gradient](excel.interior.gradient.md)** property of an **Interior** object that does not have an existing gradient fill results in a run-time error. Be aware of the **[Pattern](Excel.Interior.Pattern.md)** property of the **Interior** object before accessing the **Gradient** property.
    
If the **Pattern** property is changed from a gradient type to a non-gradient type, the **Gradient** property will populate with default values.


## Properties

- [Application](Excel.RectangularGradient.Application.md)
- [ColorStops](Excel.RectangularGradient.ColorStops.md)
- [Creator](Excel.RectangularGradient.Creator.md)
- [Parent](Excel.RectangularGradient.Parent.md)
- [RectangleBottom](Excel.RectangularGradient.RectangleBottom.md)
- [RectangleLeft](Excel.RectangularGradient.RectangleLeft.md)
- [RectangleRight](Excel.RectangularGradient.RectangleRight.md)
- [RectangleTop](Excel.RectangularGradient.RectangleTop.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]