---
title: ComboBox Control
keywords: fm20.chm5224978
f1_keywords:
- fm20.chm5224978
ms.prod: office
ms.assetid: 8a38a969-9b8c-4ba0-292c-5a3d71ce4553
ms.date: 06/08/2017
---


# ComboBox Control



Combines the features of a  **[ListBox](listbox-control.md)** and a **[TextBox](textbox-control.md)**. The user can enter a new value, as with a **[TextBox](textbox-control.md)**, or the user can select an existing value as with a **[ListBox](listbox-control.md)**.

## Remarks

If a  **[ComboBox](combobox-control.md)** is [bound](../../Glossary/glossary-vba.md#bound) to a [data source](../../Glossary/glossary-vba.md#data-source), then the  **[ComboBox](combobox-control.md)** inserts the value the user enters or selects into that data source. If a multicolumn combo box is bound, then the **BoundColumn** property determines which value is stored in the bound data source.
The list in a  **[ComboBox](combobox-control.md)** consists of rows of data. Each row can have one or more columns, which can appear with or without headings. Some applications do not support column headings, others provide only limited support.
The default property of a  **[ComboBox](combobox-control.md)** is the **Value** property.
The default event of a  **[ComboBox](combobox-control.md)** is the Change event.

 **Note**  If you want more than a single line of the list to appear at all times, you might want to use a  **[ListBox](listbox-control.md)** instead of a **[ComboBox](combobox-control.md)**. If you want to use a **[ComboBox](combobox-control.md)** and limit values to those in the list, you can set the **Style** property of the **[ComboBox](combobox-control.md)** so the control looks like a drop-down list box.


## Related Topics

[ComboBox Object](../../../api/Outlook.combobox.object.md)


