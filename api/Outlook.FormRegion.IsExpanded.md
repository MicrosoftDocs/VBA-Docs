---
title: FormRegion.IsExpanded property (Outlook)
keywords: vbaol11.chm2389
f1_keywords:
- vbaol11.chm2389
ms.prod: outlook
api_name:
- Outlook.FormRegion.IsExpanded
ms.assetid: 6b2a033c-c852-d669-d641-098f9b6c8e35
ms.date: 06/08/2017
localization_priority: Normal
---


# FormRegion.IsExpanded property (Outlook)

Returns a  **Boolean** that indicates if the form region is expanded. Read-only.


## Syntax

_expression_. `IsExpanded`

_expression_ A variable that represents a [FormRegion](Outlook.FormRegion.md) object.


## Remarks

This property applies to adjoining form regions only and is ignored for separate form regions.

Outlook always first loads a form region in an expanded state and sets  **IsExpanded** to **True**. If the initial state of the form region is to be collapsed, then Outlook immediately closes the form region, fires the **[Expanded](Outlook.FormRegion.Expanded.md)** event with the _Expand_ parameter being **False**, and sets **IsExpanded** to **False**.


## See also


[FormRegion Object](Outlook.FormRegion.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]