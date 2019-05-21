---
title: Window.ScrollRow property (Excel)
keywords: vbaxl10.chm356106
f1_keywords:
- vbaxl10.chm356106
ms.prod: excel
api_name:
- Excel.Window.ScrollRow
ms.assetid: 5fd21ea8-a173-e502-042d-57903bcd43e5
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.ScrollRow property (Excel)

Returns or sets the number of the row that appears at the top of the pane or window. Read/write **Long**.


## Syntax

_expression_.**ScrollRow**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

If the window is split, the **ScrollRow** property of the **Window** object refers to the upper-left pane. If the panes are frozen, the **ScrollRow** property of the **Window** object excludes the frozen areas.


## Example

This example moves row ten to the top of the window.

```vb
Worksheets("Sheet1").Activate 
ActiveWindow.ScrollRow = 10
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
