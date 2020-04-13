---
title: OLEFormat.ProgID property (Word)
keywords: vbawd10.chm154337302
f1_keywords:
- vbawd10.chm154337302
ms.prod: word
api_name:
- Word.OLEFormat.ProgID
ms.assetid: f3e99411-ebea-9135-e25d-66948f53e037
ms.date: 06/08/2017
localization_priority: Normal
---


# OLEFormat.ProgID property (Word)

Returns the programmatic identifier (ProgID) for the specified OLE object. Read-only  **String**.


## Syntax

_expression_. `ProgID`

_expression_ Required. A variable that represents an '[OLEFormat](Word.OLEFormat.md)' object.


## Remarks

The **ProgID** and **ClassType** properties will (by default) return the same string. However, you can change the **ClassType** property for DDE links.

For information about programmatic identifiers, see [OLE Programmatic Identifiers](overview/Word.md).


## Example

This example loops through all the floating shapes in the active document and sets all linked Microsoft Excel worksheets to be updated automatically.


```vb
For Each s In ActiveDocument.Shapes 
 If s.Type = msoLinkedOLEObject Then 
 If s.OLEFormat.ProgID = "Excel.Sheet" Then 
 s.LinkFormat.AutoUpdate = True 
 End If 
 End If 
Next
```


## See also


[OLEFormat Object](Word.OLEFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]