---
title: Window.SplitHorizontal property (Excel)
keywords: vbaxl10.chm356113
f1_keywords:
- vbaxl10.chm356113
ms.prod: excel
api_name:
- Excel.Window.SplitHorizontal
ms.assetid: 71f5aaaf-c519-dd51-410a-8f9039b11e65
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.SplitHorizontal property (Excel)

Returns or sets the location of the horizontal window split, in [points](../language/glossary/vbe-glossary.md#point). Read/write **Double**.


## Syntax

_expression_.**SplitHorizontal**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Example

This example sets the horizontal split for the active window to 216 points (3 inches).

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.SplitHorizontal = 216
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]