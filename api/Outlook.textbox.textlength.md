---
title: TextBox.TextLength Property (Outlook Forms Script)
api_name:
- Outlook.textbox.textlength
ms.assetid: 7c9ef3fe-91c4-78f5-b93d-ea5a8892b0ad
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# TextBox.TextLength Property (Outlook Forms Script)

Returns a **Long** that represents the length, in number of characters, of text in the edit region of a **[TextBox](Outlook.textbox.md)**. Read-only.


## Syntax

_expression_.**TextLength**

_expression_ A variable that represents a **TextBox** object.


## Remarks

For a multiline **TextBox**, **TextLength** includes LF (line feed) and CR (carriage return) characters.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]