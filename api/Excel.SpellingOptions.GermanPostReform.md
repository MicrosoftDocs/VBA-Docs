---
title: SpellingOptions.GermanPostReform property (Excel)
keywords: vbaxl10.chm717079
f1_keywords:
- vbaxl10.chm717079
ms.prod: excel
api_name:
- Excel.SpellingOptions.GermanPostReform
ms.assetid: 52e7c958-9122-ee2e-c5c1-335a2c2b520b
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.GermanPostReform property (Excel)

**True** to check the spelling of words by using the German post-reform rules. **False** cancels this feature. Read/write **Boolean**.


## Syntax

_expression_.**GermanPostReform**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

In this example, Microsoft Excel determines if the checking of spelling for German words is using post-reform rules and enables this feature if it's not enabled, and then notifies the user of the status.

```vb
Sub SpellingCheck() 
 
 ' Determine if spelling check for German words is using post-reform rules. 
 If Application.SpellingOptions.GermanPostReform = False Then 
 Application.SpellingOptions.GermanPostReform = True 
 MsgBox "German words will now use post-reform rules." 
 Else 
 MsgBox "German words using post-reform rules has already been set." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]