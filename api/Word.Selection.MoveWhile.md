---
title: Selection.MoveWhile method (Word)
keywords: vbawd10.chm158662768
f1_keywords:
- vbawd10.chm158662768
ms.prod: word
api_name:
- Word.Selection.MoveWhile
ms.assetid: ba35991c-2ae3-e78f-7538-c102149cf392
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.MoveWhile method (Word)

Moves the specified selection while any of the specified characters are found in the document.


## Syntax

_expression_. `MoveWhile`( `_Cset_` , `_Count_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case-sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified selection is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the specified selection is moved forward in the document, beginning at the end position. If it is a negative number, the selection is moved backward, beginning at the start position. The default value is **wdForward**.|

## Remarks

While any character in Cset is found, the specified selection is moved. The resulting  **Selection** object is positioned as an insertion point after whatever Cset characters were found. This method returns the number of characters by which the specified selection was moved, as a **Long** value. If no Cset characters are found, the selection isn't changed and the method returns 0 (zero).


## Example

This example moves the selection after consecutive tabs.


```vb
Selection.MoveWhile Cset:=vbTab, Count:=wdForward
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]