---
title: Window.GridlineColorIndex property (Excel)
keywords: vbaxl10.chm356094
f1_keywords:
- vbaxl10.chm356094
ms.prod: excel
api_name:
- Excel.Window.GridlineColorIndex
ms.assetid: c178bed5-8478-aea9-7cb4-2c7f498b533e
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.GridlineColorIndex property (Excel)

Returns or sets the gridline color as an index into the current color palette or as the following  **[xlColorIndex](Excel.XlColorIndex.md)** constant.


## Syntax

_expression_. `GridlineColorIndex`

_expression_ A variable that represents a [Window](Excel.Window.md) object.


## Remarks



| **xlColorIndex** can be the following **xlColorIndex** constant.|
| **xlColorIndexAutomatic**|

Set this property to  **xlColorIndexAutomatic** to specify the automatic color.


## Example

This example sets the gridline color in the active window to blue.


```vb
ActiveWindow.GridlineColorIndex = 5
```


## See also


[Window Object](Excel.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]