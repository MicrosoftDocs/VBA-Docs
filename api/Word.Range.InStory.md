---
title: Range.InStory method (Word)
keywords: vbawd10.chm157155453
f1_keywords:
- vbawd10.chm157155453
ms.prod: word
api_name:
- Word.Range.InStory
ms.assetid: 62452309-4d4a-5207-3e1b-28b109ca1b1e
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.InStory method (Word)

 **True** if the range to which this method is applied is in the same story as the range specified by the Range argument.


## Syntax

_expression_. `InStory`( `_Range_` )

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|Specifies the range that this method uses to determine if it is contained within the specified  **Range** object.|

## Return value

Boolean


## Remarks

A range can belong to only one story.


## Example

This example determines whether  _Range1_ and _Range2_ are in the same story. If they are, bold formatting is applied to _Range1_.


```vb
Set Range1 = Selection.Words(1) 
Set Range2 = ActiveDocument.Range(Start:=20, End:=100) 
If Range1.InStory(Range:=Range2) = True Then 
 Range1.Font.Bold = True 
End If
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]