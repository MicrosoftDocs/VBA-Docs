---
title: Options.UpdateLinksAtPrint property (Word)
keywords: vbawd10.chm162988068
f1_keywords:
- vbawd10.chm162988068
ms.prod: word
api_name:
- Word.Options.UpdateLinksAtPrint
ms.assetid: 45617b04-67ef-00f9-0161-9757fb12d1fa
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.UpdateLinksAtPrint property (Word)

 **True** if Microsoft Word updates embedded links to other files before printing a document. Read/write **Boolean**.


## Syntax

_expression_. `UpdateLinksAtPrint`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Word to update embedded links automatically before printing, and then it prints the active document.


```vb
Options.UpdateLinksAtPrint = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Update links** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.UpdateLinksAtPrint
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]