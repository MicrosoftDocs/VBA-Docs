---
title: Window.DisplayGridlines property (Excel)
keywords: vbaxl10.chm356083
f1_keywords:
- vbaxl10.chm356083
ms.prod: excel
api_name:
- Excel.Window.DisplayGridlines
ms.assetid: d4253c7f-bed2-6e58-9b04-479355f70561
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.DisplayGridlines property (Excel)

**True** if gridlines are displayed. Read/write **Boolean**.


## Syntax

_expression_.**DisplayGridlines**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

This property applies only to worksheets and macro sheets.

This property affects only displayed gridlines. Use the **[PrintGridlines](Excel.PageSetup.PrintGridlines.md)** property to control the printing of gridlines.


## Example

This example toggles the display of gridlines in the active window in Book1.xls.

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayGridlines = Not(ActiveWindow.DisplayGridlines) 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]