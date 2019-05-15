---
title: SpellingOptions.UserDict property (Excel)
keywords: vbaxl10.chm717074
f1_keywords:
- vbaxl10.chm717074
ms.prod: excel
api_name:
- Excel.SpellingOptions.UserDict
ms.assetid: 8816b44e-98e5-8829-cb6e-af4ac4040838
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.UserDict property (Excel)

Instructs Microsoft Excel to create a custom dictionary to which new words can be added when performing spelling checks on a worksheet. Read/write **String**.


## Syntax

_expression_.**UserDict**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

This example instructs Excel to create a custom dictionary called Custom1.dic in the spelling options feature and notifies the user.

```vb
Sub SpecialWord() 
 
 Application.SpellingOptions.UserDict = "Custom1.dic" 
 MsgBox "The custom dictionary is currently set to: " _ 
 & Application.SpellingOptions.UserDict 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]