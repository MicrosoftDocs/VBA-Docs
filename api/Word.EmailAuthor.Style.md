---
title: EmailAuthor.Style property (Word)
keywords: vbawd10.chm165085287
f1_keywords:
- vbawd10.chm165085287
ms.prod: word
api_name:
- Word.EmailAuthor.Style
ms.assetid: e60dadf7-affd-3bcf-e4a9-d4f083bca000
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailAuthor.Style property (Word)

Returns a  **Style** object that represents the style associated with the current email author for unsent replies, forwards, or new email messages.


## Syntax

_expression_.**Style**

_expression_ Required. A variable that represents an '[EmailAuthor](Word.EmailAuthor.md)' object.


## Example

This example returns the style associated with the current author for unsent replies, forwards, or new email messages and displays the name of the font associated with this style.


```vb
Set MyEmailStyle = _ 
 ActiveDocument.Email.CurrentEmailAuthor.Style 
Msgbox MyEmailStyle.Font.Name
```


## See also


[EmailAuthor Object](Word.EmailAuthor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]