---
title: ToggleButton control
keywords: fm20.chm2000680
f1_keywords:
- fm20.chm2000680
ms.prod: office
ms.assetid: 7ab852ce-0339-7c8b-1340-f5727cef0f56
ms.date: 11/15/2018
localization_priority: Normal
---


# ToggleButton control

Shows the selection state of an item.

## Remarks

Use a **ToggleButton** to show whether an item is selected. If a **ToggleButton** is [bound](../../Glossary/glossary-vba.md#bound) to a [data source](../../Glossary/glossary-vba.md#data-source), the **ToggleButton** shows the current value of that data source as either _Yes_/_No_, _True_/_False_, _On_/_Off_, or some other choice of two settings. 

If the user selects the **ToggleButton**, the current setting is _Yes_, _True_, or _On_; if the user does not select the **ToggleButton**, the setting is _No_, _False_, or _Off_. 

If the **ToggleButton** is bound to a data source, changing the setting changes the value of that data source. A disabled **ToggleButton** shows a value, but is dimmed and does not allow changes from the user interface.

You can also use a **ToggleButton** inside a **[Frame](frame-control.md)** to select one or more of a group of related items. For example, you can create an order form with a list of available items, with a **ToggleButton** preceding each item. The user can select a particular item by selecting the appropriate **ToggleButton**.

The default property of a **ToggleButton** is the **Value** property. The default event of a **ToggleButton** is the Click event.

## See also

- [ToggleButton object](../../../api/Outlook.togglebutton.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]