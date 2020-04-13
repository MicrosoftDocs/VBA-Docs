---
title: ComboBox.Value Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: a81934d0-50b5-aa2d-f45b-ea8b826bcea9
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.Value Property (Outlook Forms Script)

Returns or sets a **Variant** that specifies the value in the **[BoundColumn](Outlook.combobox.boundcolumn.md)** of the currently selected rows. Read/write.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **ComboBox** object.


## Remarks

Changing the contents of  **Value** does not change the value of **BoundColumn**. To add or delete entries in a **ComboBox**, you can use the  **[AddItem](Outlook.combobox.additem.md)** or **[RemoveItem](Outlook.combobox.removeitem.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]