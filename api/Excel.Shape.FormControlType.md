---
title: Shape.FormControlType property (Excel)
keywords: vbaxl10.chm636131
f1_keywords:
- vbaxl10.chm636131
ms.prod: excel
api_name:
- Excel.Shape.FormControlType
ms.assetid: a0f7d7e2-a5c0-fd71-bced-fe2866fc6d7f
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.FormControlType property (Excel)

Returns the Microsoft Excel control type. Read-only **[XlFormControl](Excel.XlFormControl.md)**.


## Syntax

_expression_.**FormControlType**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Remarks

You cannot use this property with ActiveX controls (the **[Type](Excel.Shape.Type.md)** property of the **Shape** object must return **msoFormControl**).


## Example

This example clears all the Microsoft Excel check boxes on worksheet one.

```vb
For Each s In Worksheets(1).Shapes 
 If s.Type = msoFormControl Then 
 If s.FormControlType = xlCheckBox Then _ 
 s.ControlFormat.Value = False 
 End If 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]