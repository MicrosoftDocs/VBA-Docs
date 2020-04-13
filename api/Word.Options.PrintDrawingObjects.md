---
title: Options.PrintDrawingObjects property (Word)
keywords: vbawd10.chm162988070
f1_keywords:
- vbawd10.chm162988070
ms.prod: word
api_name:
- Word.Options.PrintDrawingObjects
ms.assetid: 366ddc26-1cb0-fe48-8d54-ff9d5d3492b4
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintDrawingObjects property (Word)

 **True** if Microsoft Word prints drawing objects. Read/write **Boolean**.


## Syntax

_expression_. `PrintDrawingObjects`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Word to print drawing objects, and then it prints the active document.


```vb
Options.PrintDrawingObjects = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Drawing objects** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.PrintDrawingObjects
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]