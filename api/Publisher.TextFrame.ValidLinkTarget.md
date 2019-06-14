---
title: TextFrame.ValidLinkTarget method (Publisher)
keywords: vbapb10.chm3866662
f1_keywords:
- vbapb10.chm3866662
ms.prod: publisher
api_name:
- Publisher.TextFrame.ValidLinkTarget
ms.assetid: ee946f58-669f-7150-0f40-2dd3b857e274
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.ValidLinkTarget method (Publisher)

Determines whether the text frame of one shape can be linked to the text frame of another shape. 

Returns **True** if _LinkTarget_ is a valid target. Returns **False** if _LinkTarget_ already contains text or is already linked, or if the shape does not support attached text.


## Syntax

_expression_.**ValidLinkTarget** (_LinkTarget_)

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_LinkTarget_|Required| **Shape**|The shape with the target text frame to which you want to link the text frame returned by _expression_.|

## Return value

Boolean


## Example

This example checks to see whether the text frames for the first and second shapes on the first page of the active publication can be linked to one another. If so, the example links the two text frames.

```vb
Dim txtFrame1 As TextFrame 
Dim txtFrame2 As TextFrame 
 
With ActiveDocument.Pages(1) 
 Set txtFrame1 = .Shapes(1).TextFrame 
 Set txtFrame2 = .Shapes(2).TextFrame 
End With 
 
If txtFrame1.ValidLinkTarget(LinkTarget:=txtFrame2.Parent) = True Then 
 txtFrame1.NextLinkedTextFrame = txtFrame2 
End If
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]