---
title: Global.AutoCorrectEmail property (Word)
keywords: vbawd10.chm163119217
f1_keywords:
- vbawd10.chm163119217
ms.prod: word
api_name:
- Word.Global.AutoCorrectEmail
ms.assetid: 778d2ab6-09cb-524f-1b31-5abe467ce14c
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.AutoCorrectEmail property (Word)

Returns an  **[AutoCorrect](Word.AutoCorrect.md)** object that represents automatic corrections made to email messages.


## Syntax

_expression_. `AutoCorrectEmail`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]