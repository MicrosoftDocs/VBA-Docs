---
title: WindowState property (Excel Graph)
keywords: vbagr10.chm65932
f1_keywords:
- vbagr10.chm65932
ms.prod: excel
api_name:
- Excel.WindowState
ms.assetid: 22ce1105-6f4e-54d2-4f9a-216019462f04
ms.date: 04/12/2019
localization_priority: Normal
---


# WindowState property (Excel Graph)

Returns or sets the state of the window. Read/write **[XlWindowState](excel.xlwindowstate.md)**.

## Syntax

_expression_.**WindowState**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example maximizes the Graph application window.

```vb
myChart.Application.WindowState = xlMaximized
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]