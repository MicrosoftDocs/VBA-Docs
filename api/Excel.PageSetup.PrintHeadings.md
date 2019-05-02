---
title: PageSetup.PrintHeadings property (Excel)
keywords: vbaxl10.chm473094
f1_keywords:
- vbaxl10.chm473094
ms.prod: excel
api_name:
- Excel.PageSetup.PrintHeadings
ms.assetid: 027441c6-da40-f518-a166-adb54da02a27
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.PrintHeadings property (Excel)

**True** if row and column headings are printed with this page. Applies only to worksheets. Read/write **Boolean**.


## Syntax

_expression_.**PrintHeadings**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

The **[DisplayHeadings](Excel.Window.DisplayHeadings.md)** property of the **Window** object controls the on-screen display of headings.


## Example

This example turns off the printing of headings for Sheet1.

```vb
Worksheets("Sheet1").PageSetup.PrintHeadings = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]