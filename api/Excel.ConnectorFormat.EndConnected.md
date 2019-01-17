---
title: ConnectorFormat.EndConnected property (Excel)
keywords: vbaxl10.chm646080
f1_keywords:
- vbaxl10.chm646080
ms.prod: excel
api_name:
- Excel.ConnectorFormat.EndConnected
ms.assetid: e0831e66-f392-5044-0931-97bdab4de9c2
ms.date: 06/08/2017
localization_priority: Normal
---


# ConnectorFormat.EndConnected property (Excel)

 **msoTrue** if the end of the specified connector is connected to a shape. Read-only **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_. `EndConnected`

_expression_ A variable that represents a [ConnectorFormat](Excel.ConnectorFormat.md) object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**. Does not apply to this property.|
| **msoFalse**. The end of the specified connector is not connected to a shape.|
| **msoTriStateMixed**. Does not apply to this property.|
| **msoTriStateToggle**. Does not apply to this property.|
| **msoTrue**. The end of the specified connector is connected to a shape.|

## Example

If the end of the connector represented by shape three on  `myDocument` is connected to a shape, this example stores the connection site number in the variable `oldEndConnSite`, stores a reference to the connected shape in the object variable  `oldEndConnShape`, and then disconnects the end of the connector from the shape.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
    If .Connector Then 
        With .ConnectorFormat 
            If .EndConnected Then 
                oldEndConnSite = .EndConnectionSite 
                Set oldEndConnShape = .EndConnectedShape 
                .EndDisconnect 
            End If 
        End With 
    End If 
End With
```


## See also


[ConnectorFormat Object](Excel.ConnectorFormat.md)

