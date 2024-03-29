---
title: Height property (Excel Graph)
keywords: vbagr10.chm65659
f1_keywords:
- vbagr10.chm65659
ms.assetid: bc8f0abe-6753-a64f-4615-d0ee04a7cee4
ms.date: 04/11/2019
ms.localizationpriority: medium
---


# Height property (Excel Graph)

The height of the main application window or the object. If the window is minimized, this property is read-only and refers to the height of the icon. If the window is maximized, this property cannot be set. 

Use the **WindowState** property to determine the window state. Read/write **Double** for all objects, except for the **Chart** object, which is read/write **Variant**.

## Syntax

_expression_.**Height**

_expression_ Required. An expression that returns one of the above objects.


## Example

This example sets the height of the chart legend to 1 inch (72 points).

```vb
myChart.Legend.Height = 72
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]