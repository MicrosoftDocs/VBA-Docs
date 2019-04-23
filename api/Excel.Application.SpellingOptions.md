---
title: Application.SpellingOptions property (Excel)
keywords: vbaxl10.chm133284
f1_keywords:
- vbaxl10.chm133284
ms.prod: excel
api_name:
- Excel.Application.SpellingOptions
ms.assetid: c3d1970b-1276-9af7-88d6-e8e77bc32095
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.SpellingOptions property (Excel)

Returns a **[SpellingOptions](Excel.SpellingOptions.md)** object that represents the spelling options of the application.


## Syntax

_expression_.**SpellingOptions**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, Microsoft Excel checks the setting on the spelling options for ignoring mixed digits, and notifies the user of its status.

```vb
Sub MixedDigitCheck() 
 
 ' Determine the setting on spell checking for mixed digits. 
 If Application.SpellingOptions.IgnoreMixedDigits = True Then 
 MsgBox "The spelling options are set to ignore mixed digits." 
 Else 
 MsgBox "The spelling options are set to check for mixed digits." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]