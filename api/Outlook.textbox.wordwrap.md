---
title: TextBox.WordWrap Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: fb50b340-9fe7-17b5-4f5f-d2fdd266f37d
ms.date: 06/08/2017
localization_priority: Normal
---


# TextBox.WordWrap Property (Outlook Forms Script)

Returns or sets a **Boolean** that specifies whether the contents of a control automatically wrap at the end of a line and the control expands to fit the text. Read/write.


## Syntax

_expression_.**WordWrap**

_expression_ A variable that represents a **TextBox** object.


## Remarks

 **True** to specify that the text wraps (default), **False** to specify that the text does not.

For controls that support the  **[MultiLine](Outlook.textbox.multiline.md)** property as well as the **WordWrap** property, **WordWrap** is ignored when **MultiLine** is **False**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]