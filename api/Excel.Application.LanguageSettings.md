---
title: Application.LanguageSettings property (Excel)
keywords: vbaxl10.chm133251
f1_keywords:
- vbaxl10.chm133251
ms.prod: excel
api_name:
- Excel.Application.LanguageSettings
ms.assetid: 631879d9-f43f-4985-32d0-77bf314956eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LanguageSettings property (Excel)

Returns the  **[LanguageSettings](Office.LanguageSettings.md)** object, which contains information about the language settings in Microsoft Excel. Read-only.


## Syntax

_expression_. `LanguageSettings`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example returns the language identifier for the language you selected when you installed Microsoft Excel.


```vb
Set objLangSet = Application.LanguageSettings 
MsgBox objLangSet.LanguageID(msoLanguageIDInstall)
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
