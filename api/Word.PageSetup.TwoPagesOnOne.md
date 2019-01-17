---
title: PageSetup.TwoPagesOnOne property (Word)
keywords: vbawd10.chm158400633
f1_keywords:
- vbawd10.chm158400633
ms.prod: word
api_name:
- Word.PageSetup.TwoPagesOnOne
ms.assetid: c9d8edac-1fea-5fdb-a4e2-193920fa89d1
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.TwoPagesOnOne property (Word)

 **True** if Microsoft Word prints the specified document two pages per sheet. Read/write **Boolean**.


## Syntax

 _expression_. `TwoPagesOnOne`

 _expression_ An expression that returns a '[PageSetup](Word.PageSetup.md)' object.


## Example

This example sets Microsoft Word to print the active document two pages per sheet.


```vb
ActiveDocument.PageSetup.TwoPagesOnOne = True
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]