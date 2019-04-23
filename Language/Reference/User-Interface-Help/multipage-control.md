---
title: MultiPage control
keywords: fm20.chm5224983
f1_keywords:
- fm20.chm5224983
ms.prod: office
ms.assetid: 9361ddc3-2eaf-0a34-942b-f6cf4064625d
ms.date: 11/15/2018
localization_priority: Normal
---


# MultiPage control

Presents multiple screens of information as a single set.

## Remarks

A **MultiPage** is useful when you work with a large amount of information that can be sorted into several categories. For example, use a **MultiPage** to display information from an employment application. One page might contain personal information such as name and address; another page might list previous employers; a third page might list references. The **MultiPage** lets you visually combine related information, while keeping the entire record readily accessible.

New pages are added to the right of the currently selected page rather than adjacent to it.

> [!NOTE] 
> The **MultiPage** is a [container](../../Glossary/vbe-glossary.md#container) of a **[Pages](pages-collection-microsoft-forms.md)** collection, each of which contains one or more **[Page](page-object.md)** objects.

The default property for a **MultiPage** is the **Value** property, which returns the index of the currently active **Page** in the **Pages** collection of the **MultiPage**. The default event for a **MultiPage** is the Change event.

## See also

- [MultiPage object](../../../api/Outlook.multipage.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]