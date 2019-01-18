---
title: Application.Cursor property (Excel)
keywords: vbaxl10.chm133099
f1_keywords:
- vbaxl10.chm133099
ms.prod: excel
api_name:
- Excel.Application.Cursor
ms.assetid: 5137b89d-aba9-3e5f-b6c4-cd2264a7bd7f
ms.date: 06/08/2017
localization_priority: Priority
---


# Application.Cursor property (Excel)

Returns or sets the appearance of the mouse pointer in Microsoft Excel. Read/write  **[xlMousePointer](Excel.XlMousePointer.md)**.


## Syntax

_expression_. `Cursor`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks



| **xlMousePointer** can be one of these **xlMousePointer** constants.|
| **xlDefault**. The default pointer.|
| **xlIBeam**. The I-beam pointer.|
| **xlNorthwestArrow**. The northwest-arrow pointer.|
| **xlWait**. The hourglass pointer.|

The  **Cursor** property isn't reset automatically when the macro stops running. You should reset the pointer to **xlDefault** before your macro stops running.


## Example

This example changes the mouse pointer to an I-beam, pauses, and then changes it to the default pointer.


```vb
Sub ChangeCursor() 
 
 Application.Cursor = xlIBeam 
 For x = 1 To 1000 
 For y = 1 to 1000 
 Next y 
 Next x 
 Application.Cursor = xlDefault 
 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

