---
title: Options.DeletedTextMark property (Word)
keywords: vbawd10.chm162988090
f1_keywords:
- vbawd10.chm162988090
ms.prod: word
api_name:
- Word.Options.DeletedTextMark
ms.assetid: d1645340-5d8a-2a73-1f7f-d26278ed1950
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DeletedTextMark property (Word)

Returns or sets the format of text that is deleted while change tracking is enabled. Read/write  **WdDeletedTextMark**.


## Syntax

_expression_. `DeletedTextMark`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example applies strikethrough formatting to deleted text.


```vb
Options.DeletedTextMark = wdDeletedTextMarkStrikeThrough
```

This example returns the current status of the **Mark** option under **Deleted Text** on the **Track Changes** tab in the **Options** dialog box.




```vb
Dim lngTemp As Long 
 
lngTemp = Options.DeletedTextMark
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]