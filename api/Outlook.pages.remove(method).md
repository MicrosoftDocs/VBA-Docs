---
title: Pages.Remove Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 1b95644f-005f-e0b3-8f1e-4f125d22cad9
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.Remove Method (Outlook Forms Script)

Removes a member from a collection.


## Syntax

_expression_.**Remove**(**_varg_**)

_expression_ A variable that represents a **Pages** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|varg|Required| **Variant**|A member's position, or index, within a collection. Numeric as well as string values are acceptable. If the value is a number, the minimum value is zero, and the maximum value is one less than the number of members in the collection. If the value is a string, it must correspond to a valid member name.|

## See also


 [Pages Object](Outlook.pages(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]