---
title: Options.DiacriticColorVal property (Word)
keywords: vbawd10.chm162988453
f1_keywords:
- vbawd10.chm162988453
ms.prod: word
api_name:
- Word.Options.DiacriticColorVal
ms.assetid: bbc1c850-f4d4-7ddb-5fbf-2b9f07788a44
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DiacriticColorVal property (Word)

Returns or sets the 24-bit color to be used for diacritics in a right-to-left language document. Read/write.


## Syntax

_expression_. `DiacriticColorVal`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by the Microsoft Visual Basic **RGB** function. The value of the **UseDiffDiacColor** property must be **True** to use this property.


## Example

This example sets the color for diacritics to bright green.


```vb
If Options.UseDiffDiacColor = True Then _ 
 Options.DiacriticColorVal = wdColorBrightGreen
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]