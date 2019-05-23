---
title: Selection.Next method (Word)
keywords: vbawd10.chm158662761
f1_keywords:
- vbawd10.chm158662761
ms.prod: word
api_name:
- Word.Selection.Next
ms.assetid: 498db129-c3bd-2f9c-5897-fcfda6ce5d14
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Next method (Word)

Returns a  **Range** object that represents the next unit relative to the specified selection.


## Syntax

_expression_.**Next** (_Unit_, _Count_)

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The type of units by which to count. Can be any  **[WdUnits](Word.WdUnits.md)** constant.|
| _Count_|Optional| **Variant**|The number of units by which you want to move ahead. The default value is one.|

## Return value

Range


## Remarks

If the selection is just before the specified Unit, the selection is moved to the following unit. For example, if the selection is just before a word, the following instruction moves the selection forward to the word that follows.


```vb
Selection.Next(Unit:=wdWord, Count:=1).Select
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
