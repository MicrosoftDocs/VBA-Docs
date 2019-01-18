---
title: Application.MoveAfterReturn property (Excel)
keywords: vbaxl10.chm133168
f1_keywords:
- vbaxl10.chm133168
ms.prod: excel
api_name:
- Excel.Application.MoveAfterReturn
ms.assetid: 9cdb96d5-e28a-b30c-25de-55a807d32c25
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MoveAfterReturn property (Excel)

 **True** if the active cell will be moved as soon as the ENTER (RETURN) key is pressed. Read/write **Boolean**.


## Syntax

_expression_. `MoveAfterReturn`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks

Use the  **[MoveAfterReturnDirection](Excel.Application.MoveAfterReturnDirection.md)** property to specify the direction in which the active cell is to be moved.


## Example

This example sets the  **MoveAfterReturn** property to **True**.


```vb
Application.MoveAfterReturn = True
```


## See also


[Application Object](Excel.Application(object).md)

