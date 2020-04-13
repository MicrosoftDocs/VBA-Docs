---
title: Options.PrintOddPagesInAscendingOrder property (Word)
keywords: vbawd10.chm162988362
f1_keywords:
- vbawd10.chm162988362
ms.prod: word
api_name:
- Word.Options.PrintOddPagesInAscendingOrder
ms.assetid: c4759f97-ab6b-2df2-33b9-cf493fab1116
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintOddPagesInAscendingOrder property (Word)

 **True** if Microsoft Word prints odd pages in ascending order during manual duplex printing. Read/write **Boolean**.


## Syntax

_expression_. `PrintOddPagesInAscendingOrder`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

If the ManualDuplexPrint argument of the **[PrintOut](Word.Application.PrintOut.md)** method is **False**, this property is ignored.


## Example

This example sets Microsoft Word to print odd pages in ascending order and even pages in descending order during manual duplex printing, and then it prints the active document.


```vb
Options.PrintOddPagesInAscendingOrder = True 
Options.PrintEvenPagesInAscendingOrder = False 
ActiveDocument.PrintOut ManualDuplexPrint:=True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]