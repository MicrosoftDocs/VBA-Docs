---
title: TabStrip object (Outlook Forms Script)
keywords: olfm10.chm2000660
f1_keywords:
- olfm10.chm2000660
ms.prod: outlook
ms.assetid: 643c896a-2304-42f3-f5e9-0feee6d22364
ms.date: 06/08/2017
localization_priority: Normal
---


# TabStrip object (Outlook Forms Script)

Presents a set of related controls as a visual group.


## Remarks

You can use a **TabStrip** to view different sets of information for related controls.

A **TabStrip** is a control that contains a collection of one or more tabs.

Each **[Tab](Outlook.tab.md)** of a **TabStrip** is a separate object that users can select. Visually, a **TabStrip** also includes a client area that all the tabs in the **TabStrip** share.

By default, a **TabStrip** includes two pages, called Tab1 and Tab2. Each of these is a **Tab** object, and together they represent the **[Tabs](Outlook.tabs.md)** collection of the **TabStrip**. If you add more pages, they become part of the same **Tabs** collection.

For example, the controls might represent information about a daily schedule for a group of individuals, with each set of information corresponding to a different individual in the group. Set the title of each tab to show one individual's name. Then, you can write code that, after you click a tab, updates the controls to show information about the person identified on the tab.

The **TabStrip** is implemented as a container of a **Tabs** collection, which in turn contains a group of **Tab** objects. The **TabStrip** control does not support the **Click** event.

The default property for a **TabStrip** is the **[SelectedItem](Outlook.tabstrip.selecteditem.md)** property.


## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.tabstrip.click.md)|Occurs when the user clicks inside the control.|


## Properties

|Name|Description|
|:-----|:-----|
| [BackColor](Outlook.tabstrip.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [ClientHeight](Outlook.tabstrip.clientheight.md)|Returns a **Single** value that represents the height dimension of the display area of a [TabStrip](Outlook.tabstrip.md). Read-only.|
| [ClientLeft](Outlook.tabstrip.clientleft.md)|Returns a **Single** value that represents the location of the left edge of the display area of a **TabStrip**. Read-only.|
| [ClientTop](Outlook.tabstrip.clienttop.md)|Returns a **Single** value that represents the location of the top edge of the display area of a **TabStrip**. Read-only.|
| [ClientWidth](Outlook.tabstrip.clientwidth.md)|Returns a **Single** value that represents the width dimension of the display area of a **TabStrip**. Read-only.|
| [Enabled](Outlook.tabstrip.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.tabstrip.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [MouseIcon](Outlook.tabstrip.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.tabstrip.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [MultiRow](Outlook.tabstrip.multirow.md)|Returns or sets a **Boolean** that specifies whether the control has more than one row of tabs. Read/write.|
| [SelectedItem](Outlook.tabstrip.selecteditem.md)|Returns an **Object** that indicates the currently selected [Tab](Outlook.tab.md) object. Read-only.|
| [Style](Outlook.tabstrip.style.md)|Returns or sets an **Integer** that identifies the style of the tabs on the control. Read/write.|
| [TabFixedHeight](Outlook.tabstrip.tabfixedheight.md)|Returns or sets a **Single** that represents the height in points of the tabs on a **TabStrip**. Read/write.|
| [TabFixedWidth](Outlook.tabstrip.tabfixedwidth.md)|Returns or sets a **Single** that represents the width in points of the tabs on a **TabStrip**. Read/write.|
| [TabOrientation](Outlook.tabstrip.taborientation.md)|Returns or sets an **Integer** that specifies the location of the tabs on a **TabStrip**. Read/write.|
| [Value](Outlook.tabstrip.value.md)|Returns or sets a **Variant** that indicates the currently active tab. Read/write.|




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]