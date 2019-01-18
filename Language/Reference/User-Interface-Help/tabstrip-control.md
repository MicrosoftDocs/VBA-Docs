---
title: TabStrip control
keywords: fm20.chm2000660
f1_keywords:
- fm20.chm2000660
ms.prod: office
ms.assetid: 281a6f4a-059b-5d34-3855-f4d07b436ee4
ms.date: 11/15/2018
localization_priority: Normal
---


# TabStrip control

Presents a set of related controls as a visual group.

## Remarks

You can use a **TabStrip** to view different sets of information for related controls.

For example, the controls might represent information about a daily schedule for a group of individuals, with each set of information corresponding to a different individual in the group. Set the title of each tab to show one individual's name. You can then write code that, when you click a tab, updates the controls to show information about the person identified on the tab.

> [!NOTE] 
> The **TabStrip** is implemented as a [container](../../Glossary/vbe-glossary.md#container) of a **Tabs** collection, which in turn contains a group of **Tab** objects.

The default property for a **TabStrip** is the **SelectedItem** property. The default event for a **TabStrip** is the Change event.

## See also

- [TabStrip object](../../../api/Outlook.tabstrip.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]