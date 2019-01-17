---
title: ShapeRange.Connector property (Excel)
keywords: vbaxl10.chm640101
f1_keywords:
- vbaxl10.chm640101
ms.prod: excel
api_name:
- Excel.ShapeRange.Connector
ms.assetid: 04562f53-97a0-3f53-79de-c2c660f5a48e
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Connector property (Excel)

 **True** if the specified shape is a connector. Read-only **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_. `Connector`

 _expression_ An expression that returns a [ShapeRange](./Excel.ShapeRange.md) object.


## Example

This example deletes all connectors on  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
    For i = .Count To 1 Step -1 
        With .Item(i) 
            If .Connector Then .Delete 
        End With 
    Next 
End With
```


## See also


[ShapeRange Object](Excel.ShapeRange.md)

