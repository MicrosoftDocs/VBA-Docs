---
title: Options.DeletedTextColor property (Word)
keywords: vbawd10.chm162988093
f1_keywords:
- vbawd10.chm162988093
ms.prod: word
api_name:
- Word.Options.DeletedTextColor
ms.assetid: df77a2ad-458a-48a5-8662-6fc5ee34a003
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DeletedTextColor property (Word)

Returns or sets the color of text that is deleted while change tracking is enabled. Read/write  **WdColorIndex**.


## Syntax

_expression_. `DeletedTextColor`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

If the **DeletedTextColor** property is set to **wdByAuthor**, Word automatically assigns a unique color to each of the first eight authors who revise a document.


## Example

This example sets the color of deleted text to bright green.


```vb
Options.DeletedTextColor = wdBrightGreen
```

This example returns the current status of the **Color** option under **Deleted Text** on the **Track Changes** tab in the **Options** dialog box.




```vb
Dim lngTemp As Long 
 
lngTemp = Options.DeletedTextColor
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]