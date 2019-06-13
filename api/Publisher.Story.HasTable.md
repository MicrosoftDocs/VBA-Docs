---
title: Story.HasTable property (Publisher)
keywords: vbapb10.chm5832707
f1_keywords:
- vbapb10.chm5832707
ms.prod: publisher
api_name:
- Publisher.Story.HasTable
ms.assetid: bc4912e2-f521-c6b5-b5a6-a49952014966
ms.date: 06/14/2019
localization_priority: Normal
---


# Story.HasTable property (Publisher)

Returns **msoTrue** if the shape represents a **[Table](Publisher.Table.md)** object or **msoFalse** if the shape represents any other object type. Read-only.

<!--There is no TableFrame object, so substituted Table instead-->

## Syntax

_expression_.**HasTable**

_expression_ A variable that represents a **[Story](Publisher.Story.md)** object.


## Example

This example checks the currently selected shape to see if it is a table. If it is, the code sets the width of column one to one inch (72 points).

```vb
Sub IsTable() 
 
 With Application.Selection.ShapeRange 
 If .HasTable = msoTrue Then 
 .Table.Columns(1).Width = 72 
 End If 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]