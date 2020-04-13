---
title: Options.PrintFieldCodes property (Word)
keywords: vbawd10.chm162988064
f1_keywords:
- vbawd10.chm162988064
ms.prod: word
api_name:
- Word.Options.PrintFieldCodes
ms.assetid: f9b69b6a-2362-0370-888b-61a566803186
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintFieldCodes property (Word)

 **True** if Microsoft Word prints field codes instead of field results. Read/write **Boolean**.


## Syntax

_expression_. `PrintFieldCodes`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Word to print field codes, and then it prints the active document.


```vb
Options.PrintFieldCodes = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Field codes** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.PrintFieldCodes
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]