---
title: CellRange.Height property (Publisher)
keywords: vbapb10.chm5177348
f1_keywords:
- vbapb10.chm5177348
ms.prod: publisher
api_name:
- Publisher.CellRange.Height
ms.assetid: c4367357-5c33-7461-0cb4-a3fc392bc4bc
ms.date: 06/06/2019
localization_priority: Normal
---


# CellRange.Height property (Publisher)

Returns a **Long** that represents the height (in cells) of a table, range of cells, or page. Read-only.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[CellRange](Publisher.CellRange.md)** object.


## Remarks

The valid range for the **Height** property depends on the size of the application workspace and the position of the object within the workspace. 

For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 cells. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 cells.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]