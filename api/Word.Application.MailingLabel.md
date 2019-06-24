---
title: Application.MailingLabel property (Word)
keywords: vbawd10.chm158334994
f1_keywords:
- vbawd10.chm158334994
ms.prod: word
api_name:
- Word.Application.MailingLabel
ms.assetid: 7eba3273-4a4c-6cdf-004a-4a0d214d6127
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailingLabel property (Word)

Returns a  **[MailingLabel](Word.MailingLabel.md)** object that represents a mailing label.


## Syntax

_expression_. `MailingLabel`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


## Example

This example creates a new Avery 2160 mini-label document for a specified address.


```vb
addr = "Dave Edson" & vbCr & "123 Skye St." _ 
 & vbCr & "Our Town, WA 98004" 
Application.MailingLabel.CreateNewDocument _ 
 Name:="2160 mini", Address:=addr, ExtractAddress:=False
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]