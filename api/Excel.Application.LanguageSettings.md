---
title: Application.LanguageSettings property (Excel)
keywords: vbaxl10.chm133251
f1_keywords:
- vbaxl10.chm133251
ms.prod: excel
api_name:
- Excel.Application.LanguageSettings
ms.assetid: 631879d9-f43f-4985-32d0-77bf314956eb
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.LanguageSettings property (Excel)

Returns the **[LanguageSettings](Office.LanguageSettings.md)** object, which contains information about the language settings in Microsoft Excel. Read-only.


## Syntax

_expression_.**LanguageSettings**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example returns the language identifier for the language that you selected when you installed Microsoft Excel.

```vb
Set objLangSet = Application.LanguageSettings 
MsgBox objLangSet.LanguageID(msoLanguageIDInstall)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
