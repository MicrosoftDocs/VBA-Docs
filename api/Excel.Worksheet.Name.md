---
title: Worksheet.Name property (Excel)
keywords: vbaxl10.chm174080
f1_keywords:
- vbaxl10.chm174080
ms.prod: excel
api_name:
- Excel.Worksheet.Name
ms.assetid: 3d000cdf-5e81-8701-ca7f-bdcce006363b
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Name property (Excel)

Returns or sets a **String** value that represents the object name.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

The following code example sets the name of the active worksheet equal to today's date.

```vb
' This macro sets today's date as the name for the current sheet 
Sub NameWorksheetByDate() 
    Range("D5").Select 
    Selection.Formula = "=text(now(),""mmm dd yyyy"")" 
    Selection.Copy 
    Selection.PasteSpecial Paste:=xlValues 
    Application.CutCopyMode = False 
    Selection.Columns.AutoFit 
    ActiveSheet.Name = Range("D5").Value 
    Range("D5").Value = "" 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
