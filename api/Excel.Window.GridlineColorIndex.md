---
title: Window.GridlineColorIndex property (Excel)
keywords: vbaxl10.chm356094
f1_keywords:
- vbaxl10.chm356094
ms.prod: excel
api_name:
- Excel.Window.GridlineColorIndex
ms.assetid: c178bed5-8478-aea9-7cb4-2c7f498b533e
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.GridlineColorIndex property (Excel)

Returns or sets the gridline color as an index into the current color palette or as an **[XlColorIndex](Excel.XlColorIndex.md)** constant.


## Syntax

_expression_.**GridlineColorIndex**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

Set this property to **xlColorIndexAutomatic** to specify the automatic color.


## Example

This example sets the gridline color in the active window to blue.

```vb
ActiveWindow.GridlineColorIndex = 5
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]