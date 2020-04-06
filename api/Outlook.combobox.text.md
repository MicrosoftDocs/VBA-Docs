---
title: ComboBox.Text Property (Outlook Forms Script)
keywords: olfm10.chm2002070
f1_keywords:
- olfm10.chm2002070
ms.prod: outlook
ms.assetid: 3db98bbc-fa35-ed1f-d937-9ffeed45aed3
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.Text Property (Outlook Forms Script)

Returns or sets a  **String** that specifies text in a **[ComboBox](Outlook.combobox.md)**, changing the selected row in the control. Read/write.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a  **ComboBox** object.


## Remarks

The default value is a zero-length string ("").

You can use  **Text** to update the value of the control. If the value of **Text** matches an existing list entry, the value of the **[ListIndex](Outlook.combobox.listindex.md)** property (the index of the current row) is set to the row that matches **Text**. If the value of  **Text** does not match a row, **ListIndex** is set to -1.

When the  **Text** property of a **ComboBox** changes (such as when a user types an entry into the control), the new text is compared to the column of data specified by **[TextColumn](Outlook.combobox.textcolumn.md)**.

You cannot use  **Text** to change the value of an entry in a **ComboBox**; use the  **[Column](Outlook.combobox.column.md)** or **[List](Outlook.combobox.list.md)** property for this purpose.

The  **[ForeColor](Outlook.combobox.forecolor.md)** property determines the color of the text.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]