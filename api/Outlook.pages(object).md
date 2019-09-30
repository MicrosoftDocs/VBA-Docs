---
title: Pages object (Outlook Forms Script)
keywords: olfm10.chm0
f1_keywords:
- olfm10.chm0
ms.prod: outlook
ms.assetid: 20a5339d-1dc7-9b61-d725-d13db72c5f65
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages object (Outlook Forms Script)

Represents all the pages of a **[MultiPage](Outlook.multipage.md)**.


## Remarks

Each **Pages** collection provides the features to manage the number of pages in the collection and to identify the page that is currently in use.

The default value of the **Pages** collection identifies the current **Page** of a collection.

You can reference a **Page** by its index value. The index value reflects the ordinal position of the **Page** within the collection. The index of the first **Page** in a collection is 0; the index of the second **Page** is 1; and so on.


## Methods

|Name|Description|
|:-----|:-----|
| [Add](Outlook.Pages.add.md)|Adds a [Page](Outlook.Page.md) to a [Pages](Outlook.pages.md) collection.|
| [Clear](Outlook.Pages.clear.md)|Removes pages from the collection.|
| [Item](Outlook.Pages.item.md)|Returns a member of a collection, either by position or by name.|
| [Remove](Outlook.Pages.remove.md)|Removes a member from a collection.|


## Properties

|Name|Description|
|:-----|:-----|
| [Count](Outlook.Pages.count.md)|Returns a **Long** that represents the number of objects in a collection. Read-only.|




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]