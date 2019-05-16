---
title: SpellingOptions.HebrewModes property (Excel)
keywords: vbaxl10.chm717083
f1_keywords:
- vbaxl10.chm717083
ms.prod: excel
api_name:
- Excel.SpellingOptions.HebrewModes
ms.assetid: b8ecfa29-7ec4-180b-fb37-6876ab6c0cc7
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.HebrewModes property (Excel)

Returns or sets the mode for the Hebrew spelling checker. Read/write **[XlHebrewModes](Excel.XlHebrewModes.md)**.


## Syntax

_expression_.**HebrewModes**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Remarks

A legitimate Hebrew word can be a basic dictionary entry or any inflection.


## Example

In this example, Microsoft Excel determines the setting for the Hebrew spelling mode and notifies the user.

```vb
Sub CheckHebrewMode() 
 
 ' Determine the Hebrew spelling mode setting and notify user. 
 Select Case Application.SpellingOptions.HebrewModes 
 Case xlHebrewFullScript 
 MsgBox "The Hebrew spelling mode setting is Full Script." 
 Case xlHebrewMixedAuthorizedScript 
 MsgBox "The Hebrew spelling mode setting is Mixed Authorized Script." 
 Case xlHebrewMixedScript 
 MsgBox "The Hebrew spelling mode setting is Mixed Script." 
 Case xlHebrewPartialScript 
 MsgBox "The Hebrew spelling mode setting is Partial Script." 
 End Select 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]