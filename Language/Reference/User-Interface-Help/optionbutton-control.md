---
title: OptionButton control
keywords: fm20.chm5224984
f1_keywords:
- fm20.chm5224984
ms.prod: office
ms.assetid: 39ce3eb0-ecf1-4f1e-dbcb-a66d7d341615
ms.date: 11/15/2018
localization_priority: Normal
---


# OptionButton control

Shows the selection status of one item in a [group](../../Glossary/glossary-vba.md#control-group) of choices.

## Remarks

Use an **OptionButton** to show whether a single item in a group is selected. Note that each **OptionButton** in a **[Frame](frame-control.md)** is mutually exclusive.

If an **OptionButton** is [bound](../../Glossary/glossary-vba.md#bound) to a [data source](../../Glossary/glossary-vba.md#data-source), the **OptionButton** can show the value of that data source as either _Yes_/_No_, _True_/_False_, or _On_/_Off_. 

If the user selects the **OptionButton**, the current setting is _Yes_, _True_, or _On_; if the user does not select the **OptionButton**, the setting is _No_, _False_, or _Off_. For example, an **OptionButton** in an inventory-tracking application might show whether an item is discontinued. If the **OptionButton** is bound to a data source, changing the settings changes the value of that data source. A disabled **OptionButton** is dimmed and does not show a value.

Depending on the value of the **[TripleState](triplestate-property.md)** property, an **OptionButton** can also have a [null](../../Glossary/vbe-glossary.md#null) value.

You can also use an **OptionButton** inside a group box to select one or more of a group of related items. For example, you can create an order form with a list of available items, with an **OptionButton** preceding each item. The user can select a particular item by checking the corresponding **OptionButton**.

The default property for an **OptionButton** is the **Value** property. The default event for an **OptionButton** is the Click event.

## See also

- [OptionButton object](../../../api/Outlook.optionbutton.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]