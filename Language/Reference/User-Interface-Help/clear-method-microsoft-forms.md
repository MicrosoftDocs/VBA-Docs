---
title: Clear method (Microsoft Forms)
keywords: fm20.chm5224955
f1_keywords:
- fm20.chm5224955
ms.prod: office
ms.assetid: c0fe2f8c-1af1-6977-e794-38f9fa40deac
ms.date: 11/15/2018
localization_priority: Normal
---


# Clear method (Microsoft Forms)

Removes all objects from an object or [collection](../../Glossary/vbe-glossary.md#collection).

## Syntax

_object_. **Clear**

The **Clear** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

For a **[MultiPage](multipage-control.md)** or **[TabStrip](tabstrip-control.md)**, the **Clear** method deletes individual pages or tabs.

For a **[ListBox](listbox-control.md)** or **[ComboBox](combobox-control.md)**, **Clear** removes all entries in the list.

For a **Controls** collection, **Clear** deletes controls that were created at [run time](../../Glossary/vbe-glossary.md#run-time) with the **Add** method. Using **Clear** on controls created at [design time](../../Glossary/vbe-glossary.md#design-time) causes an error.

If the control is bound to data, the **Clear** method fails.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]