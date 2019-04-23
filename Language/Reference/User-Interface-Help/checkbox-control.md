---
title: CheckBox control
keywords: fm20.chm5224977
f1_keywords:
- fm20.chm5224977
ms.prod: office
ms.assetid: 24d90604-51ec-7f7d-e679-52391b2c27c0
ms.date: 11/15/2018
localization_priority: Normal
---


# CheckBox control

Displays the selection state of an item.

## Remarks

Use a **CheckBox** to give the user a choice between two values such as _Yes_/_No_, _True_/_False_, or _On_/_Off_.

When the user selects a **CheckBox**, it displays a special mark (such as an X) and its current setting is _Yes_, _True_, or _On_; if the user does not select the **CheckBox**, it is empty and its setting is _No_, _False_, or _Off_. Depending on the value of the **[TripleState](triplestate-property.md)** property, a **CheckBox** can also have a [null](../../Glossary/vbe-glossary.md#null) value.

If a **CheckBox** is [bound](../../Glossary/glossary-vba.md#bound) to a [data source](../../Glossary/glossary-vba.md#data-source), changing the setting changes the value of that source. A disabled **CheckBox** shows the current value, but is dimmed and does not allow changes to the value from the user interface.

You can also use check boxes inside a group box to select one or more of a group of related items. For example, you can create an order form that contains a list of available items, with a **CheckBox** preceding each item. The user can select a particular item or items by checking the corresponding **CheckBox**.

The default property of a **CheckBox** is the **Value** property. The default event of a **CheckBox** is the Click event.

> [!NOTE] 
> The **[ListBox](listbox-control.md)** also lets you put a check mark by selected options. Depending on your application, you can use the **ListBox** instead of using a group of **CheckBox** controls.

## See also

- [CheckBox object](../../../api/Outlook.checkbox.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]