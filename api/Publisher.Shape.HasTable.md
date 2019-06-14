---
title: Shape.HasTable property (Publisher)
keywords: vbapb10.chm2228321
f1_keywords:
- vbapb10.chm2228321
ms.prod: publisher
api_name:
- Publisher.Shape.HasTable
ms.assetid: 6f544d9c-00a4-3047-fbfb-6f1835bbe2c6
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.HasTable property (Publisher)

Returns **msoTrue** if the shape represents a **[Table](Publisher.Table.md)** object, or **msoFalse** if the shape represents any other object type. Read-only.

<!--There is no TableFrame object, so substituted Table instead-->

## Syntax

_expression_.**HasTable**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Remarks

The **HasTable** property value can be one of the **[MsoTriState](office.msotristate.md)** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**| The shapes in the range do not represent a **Table** object.|
| **msoTriStateMixed**|Indicates a combination of **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The shapes in the range represent a **Table** object.|

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