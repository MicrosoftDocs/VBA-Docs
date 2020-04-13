---
title: Options.PrintBackground property (Word)
keywords: vbawd10.chm162988069
f1_keywords:
- vbawd10.chm162988069
ms.prod: word
api_name:
- Word.Options.PrintBackground
ms.assetid: 3e51bfb2-63b1-d072-2a63-f3a417ffdba5
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintBackground property (Word)

 **True** if Microsoft Word prints in the background. Read/write **Boolean**.


## Syntax

_expression_. `PrintBackground`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Word to print documents in the background and then prints the active document.


```vb
Options.PrintBackground = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Background printing** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.PrintBackground
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]