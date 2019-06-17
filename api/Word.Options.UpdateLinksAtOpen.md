---
title: Options.UpdateLinksAtOpen property (Word)
keywords: vbawd10.chm162988055
f1_keywords:
- vbawd10.chm162988055
ms.prod: word
api_name:
- Word.Options.UpdateLinksAtOpen
ms.assetid: 089777c6-0bad-1fa6-4ae6-b77499c1c5a8
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.UpdateLinksAtOpen property (Word)

 **True** if Microsoft Word automatically updates all embedded OLE links in a document when it is opened. Read/write **Boolean**.


## Syntax

_expression_. `UpdateLinksAtOpen`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Word to update embedded OLE links when it opens files.


```vb
Options.UpdateLinksAtOpen = True
```

This example returns the current status of the  **Update automatic links at Open** option on the **General** tab in the **Options** dialog box.




```vb
temp = Options.UpdateLinksAtOpen
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]