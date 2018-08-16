---
title: Page Object (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 836941c3-c768-151a-65a5-41c71493033a
ms.date: 06/08/2017
---


# Page Object (Outlook Forms Script)

Represents one page of a  **[MultiPage](Outlook.multipage.md)** or a single member of a **[Pages](Outlook.pages(object).md)** collection.


## Remarks

Each  **Page** object contains its own set of controls and does not necessarily rely on other pages in the collection for information. A **Page** inherits some properties from its container; the value of each inherited property is set by the container.

You can reference a  **Page** by its index value. The index value reflects the ordinal position of the **Page** within the collection. The index of the first **Page** in a collection is 0; the index of the second **Page** is 1; and so on.

The default name for the first  **Page** is Page1. The default name for the second **Page** is Page2.


