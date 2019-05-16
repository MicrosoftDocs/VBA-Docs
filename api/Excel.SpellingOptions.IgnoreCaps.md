---
title: SpellingOptions.IgnoreCaps property (Excel)
keywords: vbaxl10.chm717075
f1_keywords:
- vbaxl10.chm717075
ms.prod: excel
api_name:
- Excel.SpellingOptions.IgnoreCaps
ms.assetid: 185b79d8-9c46-3b17-d2ee-e2544e2dce22
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.IgnoreCaps property (Excel)

**False** instructs Microsoft Excel to check for uppercase words; **True** instructs Excel to ignore words in uppercase when using the spelling checker. Read/write **Boolean**.


## Syntax

_expression_.**IgnoreCaps**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

In this example, Excel determines what the setting is for checking the spelling of uppercase words and notifies the user.

```vb
Sub SpellingOptionsCheck() 
 
 If Application.SpellingOptions.IgnoreCaps = True Then 
 MsgBox "Spelling options for checking uppercase words is disabled." 
 Else 
 MsgBox "Spelling options for checking uppercase words is enabled." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]