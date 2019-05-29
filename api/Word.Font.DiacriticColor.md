---
title: Font.DiacriticColor property (Word)
keywords: vbawd10.chm156369061
f1_keywords:
- vbawd10.chm156369061
ms.prod: word
api_name:
- Word.Font.DiacriticColor
ms.assetid: cae2bd1b-3ecb-48a4-0ba8-6273d1cd75d8
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.DiacriticColor property (Word)

Returns or sets the 24-bit color to be used for diacritics for the specified  **Font** object. Read/write.


## Syntax

_expression_. `DiacriticColor`

_expression_ Required. A variable that represents a **[Font](Word.Font.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function. The value of the **UseDiffDiacColor** property must be **True** to use this property.


## Example

This example sets the color for diacritics to blue in the current selection.


```vb
If Options.UseDiffDiacColor = True Then _ 
 Selection.Font.DiacriticColor = wdColorBlue
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]