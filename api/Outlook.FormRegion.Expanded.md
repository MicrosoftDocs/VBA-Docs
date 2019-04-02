---
title: FormRegion.Expanded event (Outlook)
keywords: vbaol11.chm2403
f1_keywords:
- vbaol11.chm2403
ms.prod: outlook
api_name:
- Outlook.FormRegion.Expanded
ms.assetid: 9d95c069-6096-6a84-f5b8-a5eeee61fde4
ms.date: 06/08/2017
localization_priority: Normal
---


# FormRegion.Expanded event (Outlook)

Occurs when the form region expands or collapses


## Syntax

_expression_. `Expanded`( `_Expand_` )

_expression_ A variable that represents a [FormRegion](Outlook.FormRegion.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Expand_|Required| **Boolean**| **True** if the form region is expanding; **False** if the form region is collapsing.|

## Remarks

This event cannot be cancelled.

Outlook always first loads a form region in an expanded state and sets  **[IsExpanded](Outlook.FormRegion.IsExpanded.md)** to **True**. If the initial state of the form region is to be collapsed, then Outlook immediately closes the form region, fires the **Expanded** event with the _Expand_ parameter being **False**, and sets **IsExpanded** to **False**.


## See also


[FormRegion Object](Outlook.FormRegion.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]