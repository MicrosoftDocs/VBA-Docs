---
title: TextRange.Move method (Publisher)
keywords: vbapb10.chm5308422
f1_keywords:
- vbapb10.chm5308422
ms.prod: publisher
api_name:
- Publisher.TextRange.Move
ms.assetid: a51b4153-2ac5-2293-d2a0-d4a3786268d7
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Move method (Publisher)

Collapses the specified range to its start position or end position, and then moves the collapsed object by the specified number of units. This method returns a **Long** that represents the number of units by which the object was actually moved, or it returns 0 (zero) if the move was unsuccessful.


## Syntax

_expression_.**Move** (_Unit_, _Size_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Unit_|Required| **[PbTextUnit](publisher.pbtextunit.md)**|The unit by which the collapsed range or selection is to be moved. Can be one of the **PbTextUnit** constants declared in the Microsoft Publisher type library.|
|_Size_|Required| **Long**|The number of units by which the specified range or selection is to be moved.<br/><br/>If _Size_ is a positive number, the object is collapsed to its end position and moved forward in the document by the specified number of units. If _Size_ is a negative number, the object is collapsed to its start position and moved backward by the specified number of units.<br/><br/>You can also control the collapse direction by using the **[Collapse](Publisher.TextRange.Collapse.md)** method before using the **Move** method.|

## Return value

Long


## Example

This example collapses the specified range and inserts a new sentence at the beginning of the range.

```vb
Sub MoveText() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Words(Start:=1, Length:=5) 
 With rngText 
 .Move Unit:=pbTextUnitParagraph, Size:=-1 
 .Text = "This adds new text to the beginning of the range. " 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]