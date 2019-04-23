---
title: ListBox.Text Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 8001cbd2-b00c-7a91-9ee6-d367ff94868b
ms.date: 06/08/2017
localization_priority: Normal
---


# ListBox.Text Property (Outlook Forms Script)

Returns or sets a  **String** that specifies text in a **[ListBox](Outlook.listbox.md)**, changing the selected row in the control. Read/write.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a  **ListBox** object.


## Remarks

The default value is a zero-length string ("").

The value of  **Text** must match an existing list entry. Specifying a value that does not match an existing list entry causes an error.

You cannot use  **Text** to change the value of an entry in a **ListBox**; use the  **[Column](Outlook.listbox.column.md)** or **[List](Outlook.listbox.list.md)** property for this purpose.

The  **[ForeColor](Outlook.listbox.forecolor.md)** property determines the color of the text.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]