---
title: Options.SaveNormalPrompt property (Word)
keywords: vbawd10.chm162988076
f1_keywords:
- vbawd10.chm162988076
ms.prod: word
api_name:
- Word.Options.SaveNormalPrompt
ms.assetid: bc58327f-d35e-70ae-ae53-0c312d3bbc0b
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.SaveNormalPrompt property (Word)

 **True** if Microsoft Word prompts the user for confirmation to save changes to the Normal template before it closes. Read/write **Boolean**.


## Syntax

_expression_. `SaveNormalPrompt`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

 **False** if Word automatically saves changes to the Normal template before it closes.


## Example

This example sets Word to save the Normal template automatically before closing, and then it quits.


```vb
Options.SaveNormalPrompt = False 
Application.Quit
```

This example returns the current status of the  **Prompt to save Normal template** option on the **Save** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.SaveNormalPrompt
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]