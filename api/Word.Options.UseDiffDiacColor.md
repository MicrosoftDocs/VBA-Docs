---
title: Options.UseDiffDiacColor property (Word)
keywords: vbawd10.chm162988452
f1_keywords:
- vbawd10.chm162988452
ms.prod: word
api_name:
- Word.Options.UseDiffDiacColor
ms.assetid: fdcf3a48-bd12-aefe-198a-81541007a596
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.UseDiffDiacColor property (Word)

 **True** if you can set the color of diacritics in the current document. Read/write **Boolean**.


## Syntax

_expression_. `UseDiffDiacColor`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example checks the  **UseDiffDiacColor** property before setting the color of diacritics in the current selection.


```vb
If Options.UseDiffDiacColor = True Then _ 
 Selection.Font.DiacriticColor = wdColorBlue
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]