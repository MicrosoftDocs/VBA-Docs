---
title: OptionButton Control
keywords: fm20.chm5224984
f1_keywords:
- fm20.chm5224984
ms.prod: office
ms.assetid: 39ce3eb0-ecf1-4f1e-dbcb-a66d7d341615
ms.date: 06/08/2017
---


# OptionButton Control



Shows the selection status of one item in a [group](../../Glossary/glossary-vba.md#control-group) of choices.

## Remarks

Use an  **[OptionButton](optionbutton-control.md)** to show whether a single item in a group is selected. Note that each **[OptionButton](optionbutton-control.md)** in a **[Frame](frame-control.md)** is mutually exclusive.
If an  **[OptionButton](optionbutton-control.md)** is [bound](../../Glossary/glossary-vba.md#bound) to a [data source](../../Glossary/glossary-vba.md#data-source), the  **[OptionButton](optionbutton-control.md)** can show the value of that data source as either _Yes_ / _No_, _True_ / _False_, or _On_ / _Off_. If the user selects the **[OptionButton](optionbutton-control.md)**, the current setting is _Yes_, _True_, or _On_; if the user does not select the **[OptionButton](optionbutton-control.md)**, the setting is _No_, _False_, or _Off_. For example, an **[OptionButton](optionbutton-control.md)** in an inventory-tracking application might show whether an item is discontinued. If the **[OptionButton](optionbutton-control.md)** is bound to a data source, then changing the settings changes the value of that data source. A disabled **[OptionButton](optionbutton-control.md)** is dimmed and does not show a value.
Depending on the value of the  **TripleState** property, an **[OptionButton](optionbutton-control.md)** can also have a [null](../../Glossary/vbe-glossary.md#null) value.
You can also use an  **[OptionButton](optionbutton-control.md)** inside a group box to select one or more of a group of related items. For example, you can create an order form with a list of available items, with an **[OptionButton](optionbutton-control.md)** preceding each item. The user can select a particular item by checking the corresponding **[OptionButton](optionbutton-control.md)**.
The default property for an  **[OptionButton](optionbutton-control.md)** is the **Value** property.
The default event for an  **[OptionButton](optionbutton-control.md)** is the Click event.

## Related Topics

[ OptionButton Object](../../../api/Outlook.optionbutton.object.md)


