---
title: ControlFormat.LinkedCell property (Excel)
keywords: vbaxl10.chm630079
f1_keywords:
- vbaxl10.chm630079
ms.prod: excel
api_name:
- Excel.ControlFormat.LinkedCell
ms.assetid: 398f46f0-593a-6020-6832-5aebe8c8cd68
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.LinkedCell property (Excel)

Returns or sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell. Read/write **String**.


## Syntax

_expression_.**LinkedCell**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Remarks

You cannot use this property with multiselect list boxes.


## Example

This example adds a check box to worksheet one and links the check box value to cell A1.

```vb
With Worksheets(1) 
 Set cb = .Shapes.AddFormControl(xlCheckBox, 10, 10, 100, 10) 
 cb.ControlFormat.LinkedCell = "A1" 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
