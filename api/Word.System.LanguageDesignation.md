---
title: System.LanguageDesignation property (Word)
keywords: vbawd10.chm154468358
f1_keywords:
- vbawd10.chm154468358
ms.prod: word
api_name:
- Word.System.LanguageDesignation
ms.assetid: c2cf7b97-262d-1b41-3d2e-58d93c243e4e
ms.date: 06/08/2017
localization_priority: Normal
---


# System.LanguageDesignation property (Word)

Returns the designated language of the system software. Read-only  **String**.


## Syntax

_expression_. `LanguageDesignation`

 _expression_ An expression that returns a '[System](Word.System.md)' object.


## Example

This example displays "U.S. English" if the **LanguageDesignation** property returns "English (US)".


```vb
If System.LanguageDesignation = "English (US)" Then _ 
 MsgBox "U.S. English"
```


## See also


[System Object](Word.System.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]