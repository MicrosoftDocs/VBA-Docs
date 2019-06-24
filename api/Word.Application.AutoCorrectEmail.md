---
title: Application.AutoCorrectEmail property (Word)
keywords: vbawd10.chm158335432
f1_keywords:
- vbawd10.chm158335432
ms.prod: word
api_name:
- Word.Application.AutoCorrectEmail
ms.assetid: 20e94c20-ead7-f16f-b70f-c37d9f34a59e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AutoCorrectEmail property (Word)

Returns an  **[AutoCorrect](Word.AutoCorrect.md)** object that represents automatic corrections made to email messages.


## Syntax

_expression_. `AutoCorrectEmail`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


## Example

This example adds AutoCorrect entries for email messages. After this code runs, every instance of "allways," "hte," and "hwen" that's typed in an email message will be replaced with "always," "the," and "when," respectively.


```vb
Sub AutoCorrectEMailAddress() 
 With Application.AutoCorrectEmail 
 .Entries.Add Name:="allways", Value:="always" 
 .Entries.Add Name:="hte", Value:="the" 
 .Entries.Add Name:="hwen", Value:="when" 
 End With 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]