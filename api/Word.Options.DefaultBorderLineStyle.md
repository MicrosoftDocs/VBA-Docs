---
title: Options.DefaultBorderLineStyle property (Word)
keywords: vbawd10.chm162988307
f1_keywords:
- vbawd10.chm162988307
ms.prod: word
api_name:
- Word.Options.DefaultBorderLineStyle
ms.assetid: 677ffe8a-ca89-fd4e-158e-158bd4c98f0c
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultBorderLineStyle property (Word)

Returns or sets the default border line style. Read/write  **WdLineStyle**.


## Syntax

_expression_. `DefaultBorderLineStyle`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the default line style to double.


```vb
Options.DefaultBorderLineStyle = wdLineStyleDouble
```

This example returns the current default line style.




```vb
Dim lngTemp As Long 
 
lngTemp= Options.DefaultBorderLineStyle
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]