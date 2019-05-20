---
title: Window.ScrollColumn property (Excel)
keywords: vbaxl10.chm356105
f1_keywords:
- vbaxl10.chm356105
ms.prod: excel
api_name:
- Excel.Window.ScrollColumn
ms.assetid: 3068b3f9-0e5e-b841-4241-7f0c060a5c25
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.ScrollColumn property (Excel)

Returns or sets the number of the leftmost column in the pane or window. Read/write **Long**.


## Syntax

_expression_.**ScrollColumn**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

If the window is split, the **ScrollColumn** property of the **[Window](Excel.Window.md)** object refers to the upper-left pane. If the panes are frozen, the **ScrollColumn** property of the **Window** object excludes the frozen areas.


## Example

This example moves column three so that it's the leftmost column in the window.

```vb
Worksheets("Sheet1").Activate 
ActiveWindow.ScrollColumn = 3
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
