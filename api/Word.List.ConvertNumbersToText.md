---
title: List.ConvertNumbersToText method (Word)
keywords: vbawd10.chm160563301
f1_keywords:
- vbawd10.chm160563301
ms.prod: word
api_name:
- Word.List.ConvertNumbersToText
ms.assetid: 302fc63e-626c-fb16-0514-25a2d6381363
ms.date: 06/08/2017
localization_priority: Normal
---


# List.ConvertNumbersToText method (Word)

Changes the list numbers and LISTNUM fields in the specified  **List** object.


## Syntax

 _expression_. `ConvertNumbersToText`

 _expression_ A variable that represents a '[List](Word.List.md)' object.


## Example

This example converts the numbers in the first list to text.


```vb
ActiveDocument.Lists(1).ConvertNumbersToText
```


## See also


[List Object](Word.List.md)

