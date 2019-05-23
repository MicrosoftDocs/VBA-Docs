---
title: Range.Next method (Word)
keywords: vbawd10.chm157155433
f1_keywords:
- vbawd10.chm157155433
ms.prod: word
api_name:
- Word.Range.Next
ms.assetid: 8d3a295d-543c-7e17-337d-b4fdfeda96e6
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Next method (Word)

Returns a  **Range** object that represents the specified unit relative to the specified range.


## Syntax

_expression_.**Next** (_Unit_, _Count_)

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The type of units by which to count. Can be any  **WdUnits** constant.|
| _Count_|Optional| **Variant**|The number of units by which you want to move ahead. The default value is one.|

## Return value

Range


## Remarks

If the range is just before the specified Unit, the range is moved to the following unit. For example, if the range is just before a word, the following instruction moves the selected text forward to the following word.


```vb
Selection.Range.Next(Unit:=wdWord, Count:=1).Select
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]