---
title: TextRange.Fields property (Publisher)
keywords: vbapb10.chm5308469
f1_keywords:
- vbapb10.chm5308469
ms.prod: publisher
api_name:
- Publisher.TextRange.Fields
ms.assetid: 01efbcae-b65b-68d9-20b0-6bbee31fd762
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.Fields property (Publisher)

Returns a  **Fields** object that represents all the fields in the specified text range.


## Syntax

_expression_.**Fields**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Return value

Fields


## Example

This example makes the first field in the first shape on the first page of the active publication bold.


```vb
Sub CountFields() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Fields(1).TextRange.Font.Bold = msoTrue 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]