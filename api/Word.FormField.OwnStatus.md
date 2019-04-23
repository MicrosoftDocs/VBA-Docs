---
title: FormField.OwnStatus property (Word)
keywords: vbawd10.chm153616390
f1_keywords:
- vbawd10.chm153616390
ms.prod: word
api_name:
- Word.FormField.OwnStatus
ms.assetid: 21595e18-6250-2f56-fc78-2336e4061055
ms.date: 06/08/2017
localization_priority: Normal
---


# FormField.OwnStatus property (Word)

Specifies the source of the text that's displayed in the status bar when a form field has the focus. Read/write  **Boolean**.


## Syntax

_expression_. `OwnStatus`

 _expression_ An expression that returns a '[FormField](Word.FormField.md)' object.


## Remarks

If  **True**, the text specified by the **[StatusText](Word.FormField.StatusText.md)** property is displayed. If **False**, the text of the AutoText entry specified by the **StatusText** property is displayed.


## Example

This example sets the status bar text for the form field named "Account" to the contents of the AutoText entry named "Acct."


```vb
With ActiveDocument.FormFields("Account") 
 .OwnStatus = False 
 .StatusText = "Acct" 
End With
```


## See also


[FormField Object](Word.FormField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]