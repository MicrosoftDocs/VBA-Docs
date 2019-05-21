---
title: Window.FreezePanes property (Excel)
keywords: vbaxl10.chm356092
f1_keywords:
- vbaxl10.chm356092
ms.prod: excel
api_name:
- Excel.Window.FreezePanes
ms.assetid: fd8c7b3b-4f70-72bd-68e4-a34442192a4e
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.FreezePanes property (Excel)

**True** if split panes are frozen. Read/write **Boolean**.


## Syntax

_expression_.**FreezePanes**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

It's possible for **FreezePanes** to be **True** and **[Split](Excel.Window.Split.md)** to be **False**, or vice versa.

This property applies only to worksheets and macro sheets.


## Example

This example freezes split panes in the active window in Book1.xls.

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.FreezePanes = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
