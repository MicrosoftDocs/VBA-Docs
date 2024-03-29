---
title: Window.SplitRow property (Excel)
keywords: vbaxl10.chm356114
f1_keywords:
- vbaxl10.chm356114
api_name:
- Excel.Window.SplitRow
ms.assetid: a1b900c3-4152-8701-db1f-1b576249c686
ms.date: 05/21/2019
ms.localizationpriority: medium
---


# Window.SplitRow property (Excel)

Returns or sets the row number where the window is split into panes (the number of rows above the split). Read/write **Long**.


## Syntax

_expression_.**SplitRow**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Example

This example splits the active window so that there are 10 rows above the split line.

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.SplitRow = 10
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]