---
title: Shape.MoveIntoTextFlow method (Publisher)
keywords: vbapb10.chm2228356
f1_keywords:
- vbapb10.chm2228356
ms.prod: publisher
api_name:
- Publisher.Shape.MoveIntoTextFlow
ms.assetid: d8a2af57-f974-717e-0d97-c8a3aee16f01
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.MoveIntoTextFlow method (Publisher)

Moves a given shape into the text flow defined by the **[TextRange](Publisher.TextRange.md)** object. The shape is always inserted inline at the beginning of the text flow.


## Syntax

_expression_.**MoveIntoTextFlow** (_Range_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Range_|Required| **TextRange**|The range of text before which the given shape is inserted.|

## Return value

Nothing


## Remarks

The **MoveIntoTextFlow** method fails if the shape to be moved is already inline or if it is not a valid inline shape type. Invalid inline shape types include:

- Inline shapes    
- Grouped shapes    
- HTML fragments    
- Smart objects    
- Chained text boxes
    

## Example

The following example checks if the second shape on the second page of the publication is inline, and if it is not, inserts it inline at the beginning of the text flow of the given text range. 

```vb
Dim theShape As Shape 
Dim theRange As TextRange 
 
Set theRange = ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
Set theShape = ActiveDocument.Pages(2).Shapes(2) 
 
If Not theShape.IsInline = msoTrue Then 
 theShape.MoveIntoTextFlow Range:=theRange 
End If 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]