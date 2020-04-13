---
title: TextBox.Copy Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: ffcb9cb8-0735-3f54-8302-d15ef14b2c27
ms.date: 06/08/2017
localization_priority: Normal
---


# TextBox.Copy Method (Outlook Forms Script)

Copies the contents of an object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a **TextBox** object.


## Remarks

The original content remains on the object.

The actual content that is copied depends on the object. For example, on a **[TextBox](Outlook.textbox.md)** or **[ComboBox](Outlook.combobox.md)**, it copies the currently selected text.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]