---
title: DoCmd.GoToPage method (Access)
keywords: vbaac10.chm4153
f1_keywords:
- vbaac10.chm4153
ms.prod: access
api_name:
- Access.DoCmd.GoToPage
ms.assetid: 37fe25b3-85b2-f681-acfd-96dab039e58f
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.GoToPage method (Access)

Carries out the GoToPage action in Visual Basic. 


## Syntax

_expression_.**GoToPage** (_PageNumber_, _Right_, _Down_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PageNumber_|Required|**Variant**|A numeric expression that's a valid page number for the active form. If you leave this argument blank, the focus stays on the current page. You can use the _Right_ and _Down_ arguments to display the part of the page that you want to see.|
| _Right_|Optional|**Variant**|A numeric expression that's a valid horizontal offset for the page.|
| _Down_|Optional|**Variant**|A numeric expression that's a valid vertical offset for the page.|

## Return value

Nothing


## Remarks

The units for the _Right_ and _Down_ arguments are expressed in [twips](../language/glossary/vbe-glossary.md#twip).

If you specify the _Right_ and _Down_ arguments and leave the _PageNumber_ argument blank, you must include the _PageNumber_ argument's comma. If you don't specify the _Right_ and _Down_ arguments, don't use a comma following the _PageNumber_ argument.

The **GoToPage** method of the **DoCmd** object was added to provide backwards compatibility for running the GoToPage action in Visual Basic code in Microsoft Access 95. We recommend that you use the existing **GoToPage** method of the **Form** object instead.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]