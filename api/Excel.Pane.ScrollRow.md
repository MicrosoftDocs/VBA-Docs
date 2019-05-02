---
title: Pane.ScrollRow property (Excel)
keywords: vbaxl10.chm360077
f1_keywords:
- vbaxl10.chm360077
ms.prod: excel
api_name:
- Excel.Pane.ScrollRow
ms.assetid: eb1f55b8-6726-00b6-845f-1cbf47cf6b13
ms.date: 05/03/2019
localization_priority: Normal
---


# Pane.ScrollRow property (Excel)

Returns or sets the number of the row that appears at the top of the pane or window. Read/write **Long**.


## Syntax

_expression_.**ScrollRow**

_expression_ A variable that represents a **[Pane](Excel.Pane.md)** object.


## Remarks

If the window is split, the **[ScrollRow](excel.window.scrollrow.md)** property of the **Window** object refers to the upper-left pane. 

If the panes are frozen, the **ScrollRow** property of the **Window** object excludes the frozen areas.


## Example

This example moves row ten to the top of the window.

```vb
Worksheets("Sheet1").Activate 
ActiveWindow.ScrollRow = 10
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]