---
title: ListBox.TextColumn Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: ecdd6bc6-f50e-9b6d-3c99-c1e282b3444a
ms.date: 06/08/2017
localization_priority: Normal
---


# ListBox.TextColumn Property (Outlook Forms Script)

Returns or sets a **Variant** that identifies the column in a **[ListBox](Outlook.listbox.md)** to display to the user. Read/write.


## Syntax

_expression_.**TextColumn**

_expression_ A variable that represents a **ListBox** object.


## Remarks

Values for the  **TextColumn** property range from -1 to the number of columns in the list. The **TextColumn** value for the first column is 1, the value of the second column is 2, and so on. Setting **TextColumn** to 0 displays the **[ListIndex](Outlook.listbox.listindex.md)** values. Setting **TextColumn** to -1 displays the first column that has a **[ColumnWidths](Outlook.listbox.columnwidths.md)** value greater than 0.

When the user selects a row from a **ComboBox** or **ListBox**, the column referenced by  **TextColumn** is stored in the **[Text](Outlook.listbox.text.md)** property. For example, you could set up a multicolumn **ListBox** that contains the names of holidays in one column and dates for the holidays in a second column. To present the holiday names to users, specify the first column as the **TextColumn**. To store the dates of the holidays, specify the second column as the  **[BoundColumn](Outlook.listbox.boundcolumn.md)**.

When the  **Text** property of a **ComboBox** **ComboBox** changes (such as when a user types an entry into the control), the new text is compared to the column of data specified by **TextColumn**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]