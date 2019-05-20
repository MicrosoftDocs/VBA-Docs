---
title: Window.DisplayHeadings property (Excel)
keywords: vbaxl10.chm356084
f1_keywords:
- vbaxl10.chm356084
ms.prod: excel
api_name:
- Excel.Window.DisplayHeadings
ms.assetid: 7105f3a4-2322-c796-5ca6-59ea46d2e248
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.DisplayHeadings property (Excel)

**True** if both row and column headings are displayed; **False** if no headings are displayed. Read/write **Boolean**.


## Syntax

_expression_.**DisplayHeadings**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

This property applies only to worksheets and macro sheets.

This property affects only displayed headings. Use the **[PrintHeadings](Excel.PageSetup.PrintHeadings.md)** property to control the printing of headings.


## Example

This example turns off the display of row and column headings in the active window in Book1.xls.

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayHeadings = False 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]