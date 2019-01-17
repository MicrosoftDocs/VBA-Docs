---
title: Options.AllowPixelUnits property (Word)
keywords: vbawd10.chm162988377
f1_keywords:
- vbawd10.chm162988377
ms.prod: word
api_name:
- Word.Options.AllowPixelUnits
ms.assetid: 11c2d832-e1e0-094e-df76-b6eeae4b0d36
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AllowPixelUnits property (Word)

 **True** if Microsoft Word uses pixels as the default unit of measurement for HTML features that support measurements. Read/write **Boolean**.


## Syntax

 _expression_. `AllowPixelUnits`

 _expression_ A variable that represents a '[Options](Word.Options.md)' object.


## Example

This example sets Word to allow pixels as the default unit of measurement for HTML features.


```vb
Options.AllowPixelUnits = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]