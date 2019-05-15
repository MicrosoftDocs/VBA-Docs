---
title: Shapes.AddFormControl method (Excel)
keywords: vbaxl10.chm638090
f1_keywords:
- vbaxl10.chm638090
ms.prod: excel
api_name:
- Excel.Shapes.AddFormControl
ms.assetid: c1654020-630c-b988-54f1-99a2f2a93e56
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddFormControl method (Excel)

Creates a Microsoft Excel control. Returns a **[Shape](Excel.Shape.md)** object that represents the new control.


## Syntax

_expression_.**AddFormControl** (_Type_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlFormControl](Excel.XlFormControl.md)**|The Microsoft Excel control type. You cannot create an edit box on a worksheet.|
| _Left_|Required| **Long**|The initial coordinates of the new object (in [points](../language/glossary/vbe-glossary.md#point)) relative to the upper-left corner of cell A1 on a worksheet or to the upper-left corner of a chart.|
| _Top_|Required| **Long**|The initial coordinates of the new object (in points) relative to the top of row 1 on a worksheet, or to the top of the chart area on a chart.|
| _Width_|Required| **Long**|The initial size of the new object, in points.|
| _Height_|Required| **Long**|The initial size of the new object, in points.|


## Return value

**Shape**


## Remarks

Use the **[AddOLEObject](Excel.Shapes.AddOLEObject.md)** method or the **[Add](Excel.OLEObjects.Add.md)** method of the **OLEObjects** collection to create an ActiveX control.


## Example

This example adds a list box to worksheet one and sets the fill range for the list box.

```vb
With Worksheets(1) 
 Set lb = .Shapes.AddFormControl(xlListBox, 100, 10, 100, 100) 
 lb.ControlFormat.ListFillRange = "A1:A10" 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]