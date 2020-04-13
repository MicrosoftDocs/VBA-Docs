---
title: ComboBox.BoundColumn Property (Outlook Forms Script)
keywords: olfm10.chm2000830
f1_keywords:
- olfm10.chm2000830
ms.prod: outlook
ms.assetid: 0ebc2ce0-f3f6-ce96-749c-be49343bc978
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.BoundColumn Property (Outlook Forms Script)

Returns or sets a **Variant** that identifies the source of data in a multicolumn **[ComboBox](Outlook.combobox.md)**. Read/write.


## Syntax

_expression_.**BoundColumn**

_expression_ A variable that represents a **ComboBox** object.


## Remarks

The possible values of  **BoundColumn** are 0 and 1. 0 assigns the value of the **[ListIndex](Outlook.combobox.listindex.md)** property to the control. 1 assigns the value from the specified column to the control. Columns are numbered from 1 when using this property (default).

When the user chooses a row in a multicolumn  **ComboBox**, the  **BoundColumn** property identifies which item from that row to store as the value of the control. For example, if each row contains 8 items and **BoundColumn** is 3, the system stores the information in the third column of the currently-selected row as the value of the object.

You can display one set of data to users but store different, associated values for the object by using the  **BoundColumn** and the **[TextColumn](Outlook.combobox.textcolumn.md)** properties. **TextColumn** identifies the column of data displayed in a **ComboBox**;  **BoundColumn** identifies the column of associated data values stored for the control. For example, you could set up a multicolumn **ComboBox** that contains the names of holidays in one column and dates for the holidays in a second column. To present the holiday names to users, specify the first column as the **TextColumn**. To store the dates of the holidays, specify the second column as the  **BoundColumn**.

The **ListIndex** value retrieves the number of the selected row. For example, if you want to know the row of the selected item, set **BoundColumn** to 0 to assign the number of the selected row as the value of the control. Be sure to retrieve a current value, rather than relying on a previously saved value, if you are referencing a list whose contents might change.

The **[Column](Outlook.combobox.column.md)**,  **[List](Outlook.combobox.list.md)**, and  **ListIndex** properties all use zero-based numbering. That is, the value of the first item (column or row) is zero; the value of the second item is one, and so on. This means that if **BoundColumn** is set to 3, you could access the value stored in that column using the expression `Column(2)`.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]