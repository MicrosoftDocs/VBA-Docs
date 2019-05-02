---
title: PageSetup.PrintNotes property (Excel)
keywords: vbaxl10.chm473095
f1_keywords:
- vbaxl10.chm473095
ms.prod: excel
api_name:
- Excel.PageSetup.PrintNotes
ms.assetid: 6609fe58-6015-9ae2-4cc0-107e29cd7b9d
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.PrintNotes property (Excel)

**True** if cell notes are printed as end notes with the sheet. Applies only to worksheets. Read/write **Boolean**.


## Syntax

_expression_.**PrintNotes**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

Use the **[PrintComments](excel.pagesetup.printcomments.md)** property to print comments as text boxes or end notes.


## Example

This example turns off the printing of notes.

```vb
Worksheets("Sheet1").PageSetup.PrintNotes = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]