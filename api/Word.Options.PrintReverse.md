---
title: Options.PrintReverse property (Word)
keywords: vbawd10.chm162988320
f1_keywords:
- vbawd10.chm162988320
ms.prod: word
api_name:
- Word.Options.PrintReverse
ms.assetid: bdbe8ff9-5d9b-a8b6-e479-338f4d2b67dd
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintReverse property (Word)

 **True** if Microsoft Word prints pages in reverse order. Read/write **Boolean**.


## Syntax

_expression_. `PrintReverse`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Word to print pages in reverse order, and then it prints the active document.


```vb
Options.PrintReverse = True 
ActiveDocument.PrintOut
```

This example returns the current status of the  **Reverse print order** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.PrintReverse
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]