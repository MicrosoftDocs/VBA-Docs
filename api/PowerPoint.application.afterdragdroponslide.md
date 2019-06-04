---
title: Application.AfterDragDropOnSlide event (PowerPoint)
keywords: vbapp10.chm621032
f1_keywords:
- vbapp10.chm621032
ms.assetid: 1de9f2a4-565b-152a-452a-cb0c1a135c35
ms.date: 06/05/2019
ms.prod: powerpoint
localization_priority: Normal
---


# Application.AfterDragDropOnSlide event (PowerPoint)

Occurs after an object with the clipboard format "PowerPoint Drop Trigger" has been dropped onto a slide in an open presentation.


## Syntax

_expression_.**AfterDragDropOnSlide** (_Sld_, _X_, _Y_)

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sld_|Required|**[Slide](PowerPoint.Slide.md)**|The slide that raised the event (that is, had a shape added to it).|
| _X_|Required|**Single**||
| _Y_|Required|**Single**||

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
