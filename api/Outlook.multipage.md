---
title: MultiPage object (Outlook Forms Script)
keywords: olfm10.chm2000570
f1_keywords:
- olfm10.chm2000570
ms.prod: outlook
ms.assetid: ac0fa233-81fe-8a34-4113-6907c6d8f7e2
ms.date: 06/08/2017
localization_priority: Normal
---


# MultiPage object (Outlook Forms Script)

Presents multiple screens of information as a single set.


## Remarks

A **MultiPage** is useful when you work with a large amount of information that can be sorted into several categories. For example, use a **MultiPage** to display information from an employment application. One page might contain personal information such as name and address; another page might list previous employers; a third page might list references. The **MultiPage** lets you visually combine related information, while keeping the entire record readily accessible.

New pages are added to the right of the currently selected page rather than adjacent to it.

A **MultiPage** is a control that contains a collection of one or more pages.

Each **[Page](Outlook.page.md)** of a **MultiPage** is a form that contains its own controls, and as such, can have a unique layout. Typically, the pages in a **MultiPage** have tabs so the user can select the individual pages.

By default, a **MultiPage** includes two pages, called Page1 and Page2. Each of these is a **Page** object, and together they represent the **[Pages](Outlook.pages(object).md)** collection of the **MultiPage**. If you add more pages, they become part of the same **Pages** collection.

The default property for a **MultiPage** is the **[Value](Outlook.multipage.value.md)** property, which returns the index of the currently active **Page** in the **Pages** collection of the **MultiPage**.

The **MultiPage** control does not support the **[Click](Outlook.multipage.click.md)** event.


## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.multipage.click.md)|Occurs when the user clicks inside the control.|


## Properties

|Name|Description|
|:-----|:-----|
| [BackColor](Outlook.multipage.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [Enabled](Outlook.multipage.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [ForeColor](Outlook.multipage.forecolor.md)|Returns or sets a **Long** that specifies the foreground color of an object. Read/write.|
| [MultiRow](Outlook.multipage.multirow.md)|Returns or sets a **Boolean** that specifies whether the control has more than one row of tabs. Read/write.|
| [SelectedItem](Outlook.multipage.selecteditem.md)|Returns an **Object** that indicates the currently selected [Page](Outlook.page.md) object. Read-only.|
| [Style](Outlook.multipage.style.md)|Returns or sets an **Integer** that identifies the style of the tabs on the control. Read/write.|
| [TabFixedHeight](Outlook.multipage.tabfixedheight.md)|Returns or sets a **Single** that represents the height in points of the tabs on a [MultiPage](Outlook.multipage.md). Read/write.|
| [TabFixedWidth](Outlook.multipage.tabfixedwidth.md)|Returns or sets a **Single** that represents the width in points of the tabs on a **MultiPage**. Read/write.|
| [TabOrientation](Outlook.multipage.taborientation.md)|Returns or sets an **Integer** that specifies the location of the tabs on a **MultiPage**. Read/write.|
| [Value](Outlook.multipage.value.md)|Returns or sets a **Variant** that indicates the currently active page. Read/write.|





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]