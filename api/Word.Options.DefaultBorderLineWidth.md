---
title: Options.DefaultBorderLineWidth property (Word)
keywords: vbawd10.chm162988316
f1_keywords:
- vbawd10.chm162988316
ms.prod: word
api_name:
- Word.Options.DefaultBorderLineWidth
ms.assetid: ab0ab48b-c05b-9be7-810e-2c97eb8ec2b7
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultBorderLineWidth property (Word)

Returns or sets the default line width of borders. Read/write  **WdLineWidth**.


## Syntax

_expression_. `DefaultBorderLineWidth`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example changes the default line width of borders and then adds a border around each paragraph in the selection.


```vb
Options.DefaultBorderLineWidth = wdLineWidth050pt 
Selection.Borders.Enable = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]