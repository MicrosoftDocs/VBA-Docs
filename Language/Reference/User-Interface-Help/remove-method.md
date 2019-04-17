---
title: Remove method (Microsoft Forms)
keywords: fm20.chm2000360
f1_keywords:
- fm20.chm2000360
ms.prod: office
api_name:
- Office.Remove
ms.assetid: 16ee4145-3e1e-9e44-7af1-2ecd3a92c9e3
ms.date: 11/15/2018
localization_priority: Normal
---


# Remove method (Microsoft Forms)

Removes a member from a [collection](../../Glossary/vbe-glossary.md#collection) or removes a control from a **[Frame](frame-control.md)**, **Page**, or form.

## Syntax

_object_.**Remove** (_collectionindex_)

The **Remove** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _collectionindex_|Required. A member's position, or index, within a collection. Numeric as well as string values are acceptable. If the value is a number, the minimum value is zero, and the maximum value is one less than the number of members in the collection. If the value is a string, it must correspond to a valid member name.|

## Remarks

This method deletes any control that was added at [run time](../../Glossary/vbe-glossary.md#run-time). However, attempting to delete a control that was added at [design time](../../Glossary/vbe-glossary.md#design-time) will result in an error.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]