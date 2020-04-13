---
title: Options.InsertedTextColor property (Word)
keywords: vbawd10.chm162988092
f1_keywords:
- vbawd10.chm162988092
ms.prod: word
api_name:
- Word.Options.InsertedTextColor
ms.assetid: 51f36823-b92b-53b0-5246-1531e851dc57
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.InsertedTextColor property (Word)

Returns or sets the color of text that is inserted while change tracking is enabled. Read/write  **WdColorIndex**.


## Syntax

_expression_. `InsertedTextColor`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

If the **InsertedTextColor** property is set to **wdByAuthor**, Microsoft Word automatically assigns a unique color to each of the first eight authors who revise a document.


## Example

This example sets the color of inserted text to dark red.


```vb
Options.InsertedTextColor = wdDarkRed
```

This example returns the current status of the **Color** option under **Track Changes** options on the **Track Changes** tab in the **Options** dialog box.




```vb
Dim lngColor As Long 
 
lngColor = Options.InsertedTextColor
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]