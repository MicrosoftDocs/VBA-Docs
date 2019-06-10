---
title: OLEFormat.ObjectVerbs property (Publisher)
keywords: vbapb10.chm4456453
f1_keywords:
- vbapb10.chm4456453
ms.prod: publisher
api_name:
- Publisher.OLEFormat.ObjectVerbs
ms.assetid: 887070e6-7f7d-4f65-290e-3d46bfd91d34
ms.date: 06/11/2019
localization_priority: Normal
---


# OLEFormat.ObjectVerbs property (Publisher)

Returns an **[ObjectVerbs](Publisher.ObjectVerbs.md)** collection that contains all the OLE verbs for the specified OLE object. Read-only.


## Syntax

_expression_.**ObjectVerbs**

_expression_ A variable that represents an **[OLEFormat](Publisher.OLEFormat.md)** object.


## Return value

ObjectVerbs


## Example

This example displays all the available verbs for the OLE object contained in shape one on page two in the active publication. For this example to work, shape one must be a shape that represents an OLE object.

```vb
Dim v As String 
 
With ActiveDocument.Pages(2).Shapes(1).OLEFormat 
 For Each v In .ObjectVerbs 
 MsgBox v 
 Next 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]