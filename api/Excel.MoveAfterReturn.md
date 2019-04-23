---
title: MoveAfterReturn property (Excel Graph)
keywords: vbagr10.chm65910
f1_keywords:
- vbagr10.chm65910
ms.prod: excel
api_name:
- Excel.MoveAfterReturn
ms.assetid: 82b4bce3-aed6-1b46-2c65-63dde6a30df1
ms.date: 04/11/2019
localization_priority: Normal
---


# MoveAfterReturn property (Excel Graph)

**True** if the active cell will be moved as soon as the Enter (Return) key is pressed. Read/write **Boolean**.

## Syntax

_expression_.**MoveAfterReturn**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the **MoveAfterReturn** property to **True**.

```vb
myChart.Application.MoveAfterReturn = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]