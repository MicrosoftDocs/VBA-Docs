---
title: Count property (Microsoft Forms)
keywords: fm20.chm2001000
f1_keywords:
- fm20.chm2001000
ms.prod: office
ms.assetid: 84580b94-05da-57d9-780b-e95545a5ea37
ms.date: 11/15/2018
localization_priority: Normal
---


# Count property (Microsoft Forms)

Returns the number of objects in a [collection](../../Glossary/vbe-glossary.md#collection).

## Syntax

_object_.**Count**

<br/>

The **Count** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

The **Count** property is read only.

Note that the index value for the first page or tab of a collection is zero, the value for the second page or tab is one, and so on. For example, if a **[MultiPage](multipage-control.md)** contains two pages, the indexes of the pages are 0 and 1, and the value of **Count** is 2.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]