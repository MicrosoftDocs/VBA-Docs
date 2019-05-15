---
title: SpellingOptions.IgnoreFileNames property (Excel)
keywords: vbaxl10.chm717078
f1_keywords:
- vbaxl10.chm717078
ms.prod: excel
api_name:
- Excel.SpellingOptions.IgnoreFileNames
ms.assetid: 346b454b-b501-9836-4d45-dbe551f4c2cb
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.IgnoreFileNames property (Excel)

**False** instructs Microsoft Excel to check for Internet and file addresses; **True** instructs Excel to ignore Internet and file addresses when using the spell checker. Read/write **Boolean**.


## Syntax

_expression_.**IgnoreFileNames**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

In this example, Excel determines what the setting is for checking the spelling of Internet and file addresses and notifies the user.

```vb
Sub SpellingOptionsCheck() 
 
 If Application.SpellingOptions.IgnoreFileNames = True Then 
 MsgBox "Spelling options for checking Internet and file addresses is disabled." 
 Else 
 MsgBox "Spelling options for checking Internet and file addresses is enabled." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]