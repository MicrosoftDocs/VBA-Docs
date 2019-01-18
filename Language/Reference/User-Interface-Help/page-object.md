---
title: Page object
keywords: fm20.chm2000590
f1_keywords:
- fm20.chm2000590
ms.prod: office
api_name:
- Office.Page
ms.assetid: 889faad0-d2ce-b404-a603-2a491c27df23
ms.date: 11/15/2018
localization_priority: Normal
---


# Page object

One page of a **[MultiPage](multipage-control.md)** and a single member of a **[Pages](pages-collection-microsoft-forms.md)** collection.

## Remarks

Each **Page** object contains its own set of controls and does not necessarily rely on other pages in the [collection](../../Glossary/vbe-glossary.md#collection) for information. A **Page** inherits some properties from its [container](../../Glossary/vbe-glossary.md#container); the value of each [inherited property](../../Glossary/glossary-vba.md#inherited-property) is set by the container.

A **Page** has a unique name and index value within a **Pages** collection. You can reference a **Page** by either its name or its index value. The index of the first **Page** in a collection is 0; the index of the second **Page** is 1; and so on. When two **Page** objects have the same name, you must reference each **Page** by its index value. References to the name in code will access only the first **Page** that uses the name.

The default name for the first **Page** is Page1; the default name for the second **Page** is Page2.

## See also

- [Page object](../../../api/Outlook.page.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]