---
title: Application.WindowsForPens property (Excel)
keywords: vbaxl10.chm133233
f1_keywords:
- vbaxl10.chm133233
ms.prod: excel
api_name:
- Excel.Application.WindowsForPens
ms.assetid: 798c0bd0-80f3-f6bd-a5d0-9abd88317bbc
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowsForPens property (Excel)

 **True** if the computer is running under Microsoft Windows for Pen Computing. Read-only **Boolean**.


## Syntax

_expression_. `WindowsForPens`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example shows how to limit handwriting recognition to numbers and punctuation only if Microsoft Windows for Pen Computing is running.


```vb
If Application.WindowsForPens Then 
 Application.ConstrainNumeric = True 
End If
```


## See also


[Application Object](Excel.Application(object).md)

