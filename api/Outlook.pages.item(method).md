---
title: Pages.Item Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: c2d80659-9741-115b-a78e-553e2b42f8d2
ms.date: 06/08/2019
localization_priority: Normal
---

# Pages.Item Method (Outlook Forms Script)

Returns a member of a collection, either by position or by name.


## Syntax
_expression_.**Item**(**_varg_**)

_expression_ A variable that represents a **Pages** object.


## Parameters
|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|varg|Required| **Variant**|A member's name or index within a collection.|

## Return value

An Object that corresponds to the specified member in the collection.

## Remarks

The  _varg_ can be either a **String** or an **Integer**. If it is a **String**, it must be a valid member name. If it is an **Integer**, the minimum value is 0 and the maximum value is one less than the number of items in the collection.

If an invalid index or name is specified, an error occurs.

## See also

 [Pages Object](Outlook.pages(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]