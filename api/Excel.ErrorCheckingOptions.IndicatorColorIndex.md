---
title: ErrorCheckingOptions.IndicatorColorIndex property (Excel)
keywords: vbaxl10.chm698074
f1_keywords:
- vbaxl10.chm698074
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.IndicatorColorIndex
ms.assetid: 4818c6b6-8cb9-705a-a441-e35159b4ffa8
ms.date: 06/08/2017
localization_priority: Normal
---


# ErrorCheckingOptions.IndicatorColorIndex property (Excel)

Returns or sets the color of the indicator for error checking options. Read/write  **[xlColorIndex](Excel.XlColorIndex.md)**.


## Syntax

_expression_. `IndicatorColorIndex`

_expression_ A variable that represents an [ErrorCheckingOptions](Excel.ErrorCheckingOptions.md) object.


## Remarks



| **xlColorIndex** can be one of these **xlColorIndex** constants.|
| **xlColorIndexAutomatic** The default system color.|
| **xlColorIndexNone** Not used with this property.|

You can specify a particular color for the indicator by entering the corresponding index value. You can use the  **Colors** property to return the current color palette.


## Example

In the following example, Microsoft Excel checks to see if the indicator color for error checking is set to the default system color and notifies the user accordingly.


```vb
Sub CheckIndexColor() 
 
 If Application.ErrorCheckingOptions.IndicatorColorIndex = xlColorIndexAutomatic Then 
 MsgBox "Your indicator color for error checking is set to the default system color." 
 Else 
 MsgBox "Your indicator color for error checking is not set to the default system color." 
 End If 
 
End Sub
```


## See also


[ErrorCheckingOptions Object](Excel.ErrorCheckingOptions.md)

