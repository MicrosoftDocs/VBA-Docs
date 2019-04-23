---
title: ListCount property
keywords: fm20.chm5225054
f1_keywords:
- fm20.chm5225054
ms.prod: office
api_name:
- Office.ListCount
ms.assetid: e6878930-514c-47cb-0961-bd9f5f79caff
ms.date: 11/16/2018
localization_priority: Normal
---


# ListCount property

Returns the number of list entries in a control.

## Syntax

_object_.**ListCount**

The **ListCount** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

The **ListCount** property is read-only. **ListCount** is the number of rows over which you can scroll. **ListRows** is the maximum to display at once. 

**ListCount** is always one greater than the largest value for the **ListIndex** property, because index numbers begin with 0 and the count of items begins with 1. If no item is selected, **ListCount** is 0 and **ListIndex** is -1.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]