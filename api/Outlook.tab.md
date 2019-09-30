---
title: Tab object (Outlook Forms Script)
ms.prod: outlook
ms.assetid: b5571953-0e47-a994-3e82-4e439a77afa8
ms.date: 06/08/2017
localization_priority: Normal
---


# Tab object (Outlook Forms Script)

Represents an individual member of a **[Tabs](Outlook.tabs.md)** collection.


## Remarks

Visually, a **Tab** object appears as a rectangle protruding from a larger rectangular area, or as a button adjacent to a rectangular area.

In contrast to a **[Page](Outlook.page.md)**, a **Tab** does not contain any controls. Controls that appear within the region bounded by a **[TabStrip](Outlook.tabstrip.md)** are contained on the form, as is the **TabStrip**.

You can reference a **Tab** by its index value. The index value reflects the ordinal position of the **Tab** within the collection. The index of the first **Tab** in a collection is 0; the index of the second **Tab** is 1; and so on.


## Properties

|Name|Description|
|:-----|:-----|
| [Accelerator](Outlook.tab.accelerator.md)|Returns or sets the accelerator key for the tab. Read/write.|
| [Caption](Outlook.tab.caption.md)|Returns or sets a **String** that specifies the text that appears on the tab. Read/write.|
| [ControlTipText](Outlook.tab.controltiptext.md)|Returns and sets a **String** that specifies text that appears when the user briefly holds the mouse pointer over a control without clicking. Read/write.|
| [Enabled](Outlook.tab.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [Index](Outlook.tab.index.md)|Returns or sets an **Integer** that specifies the position of a [Tab](Outlook.tab.md) object within a [Tabs](Outlook.tabs.md) collection. Read/write.|
| [Name](Outlook.tab.name.md)|Returns or sets a **String** that specifies the name of a control. Read/write.|
| [Tag](Outlook.tab.tag.md)|Returns or sets a **String** that specifies additional information about an object. Read/write.|
| [Visible](Outlook.tab.visible.md)|Returns or sets a **Boolean** that specifies whether a **Tab** is visible or hidden. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]