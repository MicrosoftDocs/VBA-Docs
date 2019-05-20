---
title: Window.View property (Excel)
keywords: vbaxl10.chm356127
f1_keywords:
- vbaxl10.chm356127
ms.prod: excel
api_name:
- Excel.Window.View
ms.assetid: 604ea4f4-8268-9939-cac3-2e082a2c4831
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.View property (Excel)

Returns or sets the view showing in the window. Read/write **[XlWindowView](Excel.XlWindowView.md)**.


## Syntax

_expression_.**View**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Example

This example switches the view in the active window to page break preview.

```vb
ActiveWindow.View = xlPageBreakPreview
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]