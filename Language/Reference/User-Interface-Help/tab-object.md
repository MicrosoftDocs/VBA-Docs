---
title: Tab object
keywords: fm20.chm2000640
f1_keywords:
- fm20.chm2000640
ms.prod: office
api_name:
- Office.Tab
ms.assetid: 3ac18010-5e79-c52b-e2ae-0fcd3acdd7b9
ms.date: 11/15/2018
localization_priority: Normal
---


# Tab object

A **Tab** is an individual member of a **[Tabs](tabs-collection-microsoft-forms.md)** collection.

## Remarks

Visually, a **Tab** object appears as a rectangle protruding from a larger rectangular area or as a button adjacent to a rectangular area.

In contrast to a **Page**, a **Tab** does not contain any controls. Controls that appear within the region bounded by a **[TabStrip](tabstrip-control.md)** are contained on the form, as is the **TabStrip**.

Each **Tab** has its own set of properties, but has no methods or events. You must use events from the appropriate **TabStrip** to initiate processing of an individual **Tab**.

Each **Tab** has a unique name and index value within the [collection](../../Glossary/vbe-glossary.md#collection). You can reference a **Tab** by either its name or its index value. The index of the first **Tab** is 0; the index of the second **Tab** is 1; and so on. When two **Tab** objects have the same name, you must reference each **Tab** by its index value. References to the name in code will access only the first **Tab** that uses the name.

## See also

- [Tab object](../../../api/Outlook.tab.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]