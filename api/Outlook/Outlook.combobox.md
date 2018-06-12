---
title: ComboBox Object (Outlook Forms Script)
keywords: olfm10.chm2000480
f1_keywords:
- olfm10.chm2000480
ms.prod: outlook
ms.assetid: 31e7c1de-ee4e-b3d9-4579-7fc6b215bad3
ms.date: 06/08/2017
---


# ComboBox Object (Outlook Forms Script)

Combines the features of a  **[ListBox](Outlook.listbox.md)** and a **[TextBox](Outlook.textbox.md)**. 


## Remarks

The user can enter a new value, as with a  **TextBox**, or the user can select an existing value as with a  **ListBox**.

If a  **ComboBox** is bound to a data source, the **ComboBox** inserts the value entered or selected by the user into that data source. If a multicolumn combo box is bound, then the **[BoundColumn](Outlook.combobox.boundcolumn.md)** property determines which value is stored in the bound data source.

The list in a  **ComboBox** consists of rows of data. Each row can have one or more columns, which can appear with or without headings. Some applications do not support column headings, others provide only limited support.

The default property of a  **ComboBox** is the **[Value](Outlook.combobox.value.md)** property.

If you want more than a single line of the list to appear at all times, you might want to use a  **ListBox** instead of a **ComboBox**. If you want to use a  **ComboBox** and limit values to those in the list, you can set the **[Style](Outlook.combobox.style.md)** property of the **ComboBox** so the control looks like a drop-down list box.


