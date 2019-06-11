---
title: PhoneticGuide.BaseText property (Publisher)
keywords: vbapb10.chm6160391
f1_keywords:
- vbapb10.chm6160391
ms.prod: publisher
api_name:
- Publisher.PhoneticGuide.BaseText
ms.assetid: e59ef54f-c650-1a3e-717b-b4b603f312c1
ms.date: 06/12/2019
localization_priority: Normal
---


# PhoneticGuide.BaseText property (Publisher)

Returns a **String** that represents the text to which the specified phonetic text applies. Read-only.


## Syntax

_expression_.**BaseText**

_expression_ A variable that represents a **[PhoneticGuide](Publisher.PhoneticGuide.md)** object.


## Return value

String


## Example

This example adds phonetic text to the selection and displays the text to which the phonetic text applies, which is the originally-selected text. This example assumes that text is selected. If no text is selected, the message box will be blank.

```vb
Sub AddPhoneticText() 
 With Selection.TextRange.Fields.AddPhoneticGuide _ 
 (Range:=Selection.TextRange, Text:="tray sheek") 
 MsgBox "The base text is " & .PhoneticGuide.BaseText 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]