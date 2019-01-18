---
title: TextFrame2.Creator property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.Creator
ms.assetid: 12c1e3ee-4c76-907a-2606-661108f8a6ae
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.Creator property (Office)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only. Long


## Syntax

_expression_. `Creator`

 _expression_ An expression that returns a [TextFrame2](Office.TextFrame2.md) object.


## Example

This example displays a message about the creator of an Excel workbook. In this example, the hexadecimal number 5843454C is equivalent to the string XCEL which indicates that this object was created in Excel.


```vb
Sub FindCreator() 
 
 Dim myObject As Excel.Workbook 
 Set myObject = ActiveWorkbook 
 If myObject.TextFrame2.Creator = &amp;h5843454c Then 
 MsgBox "This is a Microsoft Excel object." 
 Else 
 MsgBox "This is not a Microsoft Excel object." 
 End If 
 
End Sub 

```


## See also


[TextFrame2 Object](Office.TextFrame2.md)



[TextFrame2 Object Members](./overview/Library-Reference/textframe2-members-office.md)

