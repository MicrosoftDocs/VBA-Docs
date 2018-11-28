---
title: Application.MoveAfterReturnDirection property (Excel)
keywords: vbaxl10.chm133169
f1_keywords:
- vbaxl10.chm133169
ms.prod: excel
api_name:
- Excel.Application.MoveAfterReturnDirection
ms.assetid: c11d8e36-755e-c911-de44-8b630b549418
ms.date: 06/08/2017
---


# Application.MoveAfterReturnDirection property (Excel)

Returns or sets the direction in which the active cell is moved when the user presses ENTER. Read/write  **[xlDirection](Excel.XlDirection.md)**.


## Syntax

 _expression_. `MoveAfterReturnDirection`

 _expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks



| **xlDirection** can be one of these **xlDirection** constants.|
| **xlDown**|
| **xlToLeft**|
| **xlToRight**|
| **xlUp**|

If the  **[MoveAfterReturn](Excel.Application.MoveAfterReturn.md)** property is **False** , the selection doesn't move at all, regardless of how the **MoveAfterReturnDirection** property is set.


## Example

This example causes the active cell to move to the right when the user presses ENTER.


```vb
Application.MoveAfterReturn = True 
Application.MoveAfterReturnDirection = xlToRight
```


## See also


[Application Object](Excel.Application(object).md)

