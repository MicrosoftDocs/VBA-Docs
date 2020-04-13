---
title: OlkTextBox.PasswordChar property (Outlook)
keywords: vbaol11.chm1000054
f1_keywords:
- vbaol11.chm1000054
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.PasswordChar
ms.assetid: 1632642a-2948-4cc0-b086-ae454ae9a7ed
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTextBox.PasswordChar property (Outlook)

Returns or sets a **String** that specifies a placeholder character that is to be displayed repetitively as a string instead of the actual characters entered in the text box. Read/write.


## Syntax

_expression_. `PasswordChar`

_expression_ A variable that represents an [OlkTextBox](Outlook.OlkTextBox.md) object.


## Remarks

A common application of setting this property to  **True** is password entry, where you would not want to display the actual characters of the password that the user enters in the text box. The default value is an empty string.

Only one character is accepted for the value of this property. If a string of more than one character is set, only the first character will be used as the placeholder character and the rest will be ignored.


## See also


[OlkTextBox Object](Outlook.OlkTextBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]