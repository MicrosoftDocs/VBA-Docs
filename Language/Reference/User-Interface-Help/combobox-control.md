---
title: ComboBox control
keywords: fm20.chm5224978
f1_keywords:
- fm20.chm5224978
ms.prod: office
ms.assetid: 8a38a969-9b8c-4ba0-292c-5a3d71ce4553
ms.date: 11/15/2018
localization_priority: Normal
---


# ComboBox control

Combines the features of a **[ListBox](listbox-control.md)** and a **[TextBox](textbox-control.md)**. The user can enter a new value, as with a **TextBox**, or the user can select an existing value, as with a **ListBox**.

## Remarks

If a **ComboBox** is [bound](../../Glossary/glossary-vba.md#bound) to a [data source](../../Glossary/glossary-vba.md#data-source), the **ComboBox** inserts the value the user enters or selects into that data source. If a multi-column combo box is bound, the **BoundColumn** property determines which value is stored in the bound data source.

The list in a **ComboBox** consists of rows of data. Each row can have one or more columns, which can appear with or without headings. Some applications do not support column headings, others provide only limited support.

The default property of a **ComboBox** is the **Value** property. The default event of a **ComboBox** is the Change event.

> [!NOTE] 
> If you want more than a single line of the list to appear at all times, you might want to use a **ListBox** instead of a **ComboBox**. If you want to use a **ComboBox** and limit values to those in the list, you can set the **Style** property of the **ComboBox** so the control looks like a drop-down list box.

## See also

- [ComboBox object](../../../api/Outlook.combobox.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
