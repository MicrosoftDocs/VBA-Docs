---
title: Range.Phonetics property (Excel)
keywords: vbaxl10.chm144229
f1_keywords:
- vbaxl10.chm144229
ms.prod: excel
api_name:
- Excel.Range.Phonetics
ms.assetid: fdc05b76-b574-63ec-045a-42fdcfae8a9e
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Phonetics property (Excel)

Returns the **[Phonetics](Excel.Phonetics.md)** collection of the range. Read-only.


## Syntax

_expression_.**Phonetics**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example displays all the **Phonetic** objects in the active cell.

```vb
Set objPhon = ActiveCell.Phonetics 
With objPhon 
 For Each objPhonItem in objPhon 
 MsgBox "Phonetic object: " & .Text 
 Next 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]