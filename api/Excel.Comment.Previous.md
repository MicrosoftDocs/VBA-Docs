---
title: Comment.Previous method (Excel)
keywords: vbaxl10.chm516079
f1_keywords:
- vbaxl10.chm516079
ms.prod: excel
api_name:
- Excel.Comment.Previous
ms.assetid: b7854b0f-0e88-6749-2e62-6d45add8b945
ms.date: 04/23/2019
localization_priority: Normal
---


# Comment.Previous method (Excel)

Returns a **Comment** object that represents the previous comment.


## Syntax

_expression_.**Previous**

_expression_ An expression that returns a **[Comment](Excel.Comment.md)** object.


## Return value

Comment


## Remarks

This method works only on one sheet. Using this method on the first comment on a sheet returns **Null** (not the last comment on the previous sheet).


## Example

This example deletes every second comment, navigating with the **Previous** method.

> [!NOTE] 
> Test this example in a new workbook with no existing comments. To clear all the comments from a workbook, use  `Selection.SpecialCells(xlCellTypeComments).delete` in the Immediate pane.


```vb
'Sets up the comments 
For xNum = 1 To 10 
 Range("A" & xNum).AddComment 
 Range("A" & xNum).Comment.Text Text:="Comment " & xNum 
Next 
 
MsgBox "Comments created... A1:A10" 
 
'Deletes every second comment in the A1:A10 range 
For yNum = 10 To 1 Step -2 
 Range("A" & yNum).Comment.Previous.Shape.Select True 
 Selection.Delete 
Next 
 
MsgBox "Deleted every second comment"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]