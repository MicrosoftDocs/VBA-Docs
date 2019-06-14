---
title: ShapeRange.Height property (Publisher)
keywords: vbapb10.chm2293817
f1_keywords:
- vbapb10.chm2293817
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Height
ms.assetid: de6a638d-c197-a35b-130e-a9507d1b918e
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Height property (Publisher)

Returns a **Variant** that represents the height (in [points](../language/glossary/vbe-glossary.md#point)) of a specified range of shapes. Read-only.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

The valid range for the **Height** property depends on the size of the application workspace and the position of the object within the workspace. 

For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]