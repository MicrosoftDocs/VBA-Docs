---
title: Application.AutoCorrect property (Excel)
keywords: vbaxl10.chm133081
f1_keywords:
- vbaxl10.chm133081
ms.prod: excel
api_name:
- Excel.Application.AutoCorrect
ms.assetid: e339617e-e086-7324-9240-4db9cfcfcee5
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AutoCorrect property (Excel)

Returns an **[AutoCorrect](Excel.AutoCorrect(object).md)** object that represents the Microsoft Excel AutoCorrect attributes. Read-only.


## Syntax

_expression_.**AutoCorrect**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example substitutes the word Temp. for the word Temperature in the array of AutoCorrect replacements.

```vb
With Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]