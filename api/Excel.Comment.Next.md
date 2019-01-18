---
title: Comment.Next method (Excel)
keywords: vbaxl10.chm516078
f1_keywords:
- vbaxl10.chm516078
ms.prod: excel
api_name:
- Excel.Comment.Next
ms.assetid: 0331918c-056d-6adc-e232-0aeee3d9c57b
ms.date: 06/08/2017
localization_priority: Normal
---


# Comment.Next method (Excel)

Returns a  **[Comment](Excel.Comment.md)** object that represents the next comment.


## Syntax

_expression_. `Next`

 _expression_ An expression that returns a [Comment](Excel.Comment.md) object.


## Return value

Comment


## Remarks

This method works only on one sheet. Using this method on the last comment on a sheet returns  **Null** (not the next comment on the next sheet).


## Example

This example shows every second comment, navigating with the next method.


 **Note**  Please test in a new workbook with no existing comments. To clear all comments from a workbook use  `Selection.SpecialCells(xlCellTypeComments).delete` in the **Immediate Pane**.


```vb
'Sets up the comments 
For xNum = 1 To 10 
 Range("A" & xNum).AddComment 
 Range("A" & xNum).Comment.Text Text:="Comment " & xNum 
Next 
 
MsgBox "Comments created... A1:A10" 
 
'Deletes every second comment in the A1:A10 range 
For yNum = 1 To 10 Step 2 
 Range("A" & yNum).Comment.Next.Shape.Select True 
 Selection.Delete 
Next 
 
MsgBox "Deleted every second comment"
```


## See also


[Comment Object](Excel.Comment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]