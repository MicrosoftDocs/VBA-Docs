---
title: Pane.ScrollColumn property (Excel)
keywords: vbaxl10.chm360076
f1_keywords:
- vbaxl10.chm360076
ms.prod: excel
api_name:
- Excel.Pane.ScrollColumn
ms.assetid: 47165fe4-299d-8863-708f-9db22ef52ed1
ms.date: 05/03/2019
localization_priority: Normal
---


# Pane.ScrollColumn property (Excel)

Returns or sets the number of the leftmost column in the pane or window. Read/write **Long**.


## Syntax

_expression_.**ScrollColumn**

_expression_ A variable that represents a **[Pane](Excel.Pane.md)** object.


## Remarks

If the window is split, the **[ScrollColumn](excel.window.scrollcolumn.md)** property of the **Window** object refers to the upper-left pane. 

If the panes are frozen, the **ScrollColumn** property of the **Window** object excludes the frozen areas.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]