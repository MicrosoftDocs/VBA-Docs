---
title: Shape.Child Property (Excel)
keywords: vbaxl10.chm636138
f1_keywords:
- vbaxl10.chm636138
ms.prod: excel
api_name:
- Excel.Shape.Child
ms.assetid: fa3a7f15-8f55-3c7f-4d4f-5af3744fe022
ms.date: 06/08/2017
---


# Shape.Child Property (Excel)

Returns  **msoTrue** if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only **[MsoTriState](./Office.MsoTriState.md)** .


## Syntax

 _expression_. `Child`

 _expression_ A variable that represents a [Shape](./Excel.Shape.md) object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue** . Does not apply to this property.|
| **msoFalse** . If the selected shape is not a child shape.|
| **msoTriStateMixed** . If only some of the selected shapes are child shapes.|
| **msoTriStateToggle** . Does not apply to this property.|
| **msoTrue** . If the selected shape is a child shape.|

## See also


[Shape Object](Excel.Shape.md)

