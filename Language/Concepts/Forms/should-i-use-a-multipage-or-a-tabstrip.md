---
title: Should I use a MultiPage or a TabStrip?
keywords: fm20.chm5225198
f1_keywords:
- fm20.chm5225198
ms.prod: office
ms.assetid: 3da861ef-58ca-1993-d661-b20c3d337673
ms.date: 12/29/2018
localization_priority: Normal
---


# Should I use a MultiPage or a TabStrip?

If you use a single layout for data, use a **[TabStrip](../../reference/user-interface-help/tabstrip-control.md)** and map each set of data to its own **[Tab](../../reference/user-interface-help/tab-object.md)**. If you need several layouts for data, use a **[MultiPage](../../reference/user-interface-help/multipage-control.md)** and assign each layout to its own **[Page](../../reference/user-interface-help/page-object.md)**.

Unlike a **Page** of a **MultiPage**, the [client region](../../Glossary/glossary-vba.md#client-region) of a **TabStrip** is not a separate form, but a portion of the form that contains the **TabStrip**. The border of a **TabStrip** defines a region of the form that you can associate with the tabs. When you place a control in the client region of a **TabStrip**, you are adding a control to the form that contains the **TabStrip**.

## See also

- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]