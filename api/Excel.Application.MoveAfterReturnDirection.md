---
title: Application.MoveAfterReturnDirection property (Excel)
keywords: vbaxl10.chm133169
f1_keywords:
- vbaxl10.chm133169
ms.prod: excel
api_name:
- Excel.Application.MoveAfterReturnDirection
ms.assetid: c11d8e36-755e-c911-de44-8b630b549418
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.MoveAfterReturnDirection property (Excel)

Returns or sets the direction in which the active cell is moved when the user presses Enter. Read/write **[XlDirection](Excel.XlDirection.md)**.


## Syntax

_expression_.**MoveAfterReturnDirection**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

**XlDirection** can be one of these constants:

- **xlDown**
- **xlToLeft**
- **xlToRight**
- **xlUp**

If the **[MoveAfterReturn](Excel.Application.MoveAfterReturn.md)** property is **False**, the selection doesn't move at all, regardless of how the **MoveAfterReturnDirection** property is set.


## Example

This example causes the active cell to move to the right when the user presses Enter.

```vb
Application.MoveAfterReturn = True 
Application.MoveAfterReturnDirection = xlToRight
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]