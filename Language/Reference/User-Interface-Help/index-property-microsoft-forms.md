---
title: Index property (Microsoft Forms)
keywords: fm20.chm5225044
f1_keywords:
- fm20.chm5225044
ms.prod: office
ms.assetid: 304f42ff-5a38-0e84-8f9f-40e75d7fc2b2
ms.date: 11/16/2018
localization_priority: Normal
---


# Index property (Microsoft Forms)

The position of a **[Tab](tab-object.md)** object within a **[Tabs](tabs-collection-microsoft-forms.md)** collection, or a **[Page](page-object.md)** object in a **[Pages](pages-collection-microsoft-forms.md)** collection.

## Syntax

_object_.**Index** [= _Integer_ ]

The **Index** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Integer_|Optional. The index of the currently selected **Tab** object.|

## Remarks

The **Index** property specifies the order in which tabs appear. Changing the value of **Index** visually changes the order of **Pages** in a **[MultiPage](multipage-control.md)** or **Tabs** on a **[TabStrip](tabstrip-control.md)**. The index value for the first page or tab is zero, the index value of the second page or tab is one, and so on.

In a **MultiPage**, **Index** refers to a **Page** as well as the page's **Tab**. In a **TabStrip**, **Index** refers to the tab only.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]