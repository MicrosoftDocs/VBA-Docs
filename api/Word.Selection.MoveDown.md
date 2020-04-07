---
title: Selection.MoveDown method (Word)
keywords: vbawd10.chm158663159
f1_keywords:
- vbawd10.chm158663159
ms.prod: word
api_name:
- Word.Selection.MoveDown
ms.assetid: d3ea31e8-04a5-c342-24ca-c93ac1a1258e
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.MoveDown method (Word)

Moves the selection down and returns the number of units it has been moved.


## Syntax

_expression_. `MoveDown`( `_Unit_` , `_Count_` , `_Extend_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which the selection is to be moved.The default value is **wdLine**.|
| _Count_|Optional| **Variant**|The number of units the selection is to be moved. The default value is 1.|
| _Extend_|Optional| **Variant**|Can be either  **wdMove** or **wdExtend**. If **wdMove** is used, the selection is collapsed to the endpoint and moved down. If **wdExtend** is used, the selection is extended down. The default value is **wdMove**.|

## Example

This example extends the selection down one line.


```vb
Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
```

This example moves the selection down three paragraphs. If the move is successful, "Company" is inserted at the insertion point.




```vb
unitsMoved = Selection.MoveDown(Unit:=wdParagraph, _ 
 Count:=3, Extend:=wdMove) 
If unitsMoved = 3 Then Selection.Text = "Company"
```

This example displays the current line number, moves the selection down three lines, and displays the current line number again.




```vb
MsgBox "Line " & Selection.Information(wdFirstCharacterLineNumber) 
Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdMove 
MsgBox "Line " & Selection.Information(wdFirstCharacterLineNumber)
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
