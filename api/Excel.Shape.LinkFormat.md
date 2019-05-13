---
title: Shape.LinkFormat property (Excel)
keywords: vbaxl10.chm636129
f1_keywords:
- vbaxl10.chm636129
ms.prod: excel
api_name:
- Excel.Shape.LinkFormat
ms.assetid: f364d08e-aafd-1555-34ee-f0682cde7e19
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.LinkFormat property (Excel)

Returns a **[LinkFormat](Excel.LinkFormat.md)** object that contains linked OLE object properties. Read-only.


## Syntax

_expression_.**LinkFormat**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

This example updates all linked OLE objects on worksheet one.

```vb
For Each s In Worksheets(1).Shapes 
 If s.Type = msoLinkedOLEObject Then s.LinkFormat.Update 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]