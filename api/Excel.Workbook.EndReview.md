---
title: Workbook.EndReview method (Excel)
keywords: vbaxl10.chm199208
f1_keywords:
- vbaxl10.chm199208
ms.prod: excel
api_name:
- Excel.Workbook.EndReview
ms.assetid: cd4a445b-4731-43ba-e46a-f80f19ea5a17
ms.date: 06/08/2017
localization_priority: Normal
---


# Workbook.EndReview method (Excel)

Terminates a review of a file that has been sent for review using the  **[SendForReview](Excel.Workbook.SendForReview.md)** method.


## Syntax

_expression_. `EndReview`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Example

This example terminates the review of the active workbook. When executed, this procedure displays a message asking if you want to end the review. This example assumes the active workbook has been sent for review.


```vb
Sub EndWorkbookRev() 
 
 ActiveWorkbook.EndReview 
 
End Sub
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]