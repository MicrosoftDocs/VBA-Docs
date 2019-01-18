---
title: PageSetup.LinesPage property (Word)
keywords: vbawd10.chm158400636
f1_keywords:
- vbawd10.chm158400636
ms.prod: word
api_name:
- Word.PageSetup.LinesPage
ms.assetid: e063f2e4-d7de-48b4-15b0-db75ca9fb6e4
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.LinesPage property (Word)

Returns or sets the number of lines per page in the document grid. Read/write  **Single**.


## Syntax

 _expression_. `LinesPage`

 _expression_ An expression that returns a '[PageSetup](Word.PageSetup.md)' object.


## Example

This example sets the number of lines per page to 35 for the active document.


```vb
ActiveDocument.PageSetup.LinesPage = 35
```


## See also


[PageSetup Object](Word.PageSetup.md)

