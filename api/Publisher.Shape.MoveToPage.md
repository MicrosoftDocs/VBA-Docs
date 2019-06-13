---
title: Shape.MoveToPage method (Publisher)
keywords: vbapb10.chm2228376
f1_keywords:
- vbapb10.chm2228376
ms.prod: publisher
api_name:
- Publisher.Shape.MoveToPage
ms.assetid: 1893035f-6739-7480-6ba0-2ca6a42355fa
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.MoveToPage method (Publisher)

Moves a shape to the specified page.


## Syntax

_expression_.**MoveToPage** (_Page_, _Left_, _Top_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Page_ |Required| **Long**|Page to which the shape should be moved.|
|_Left_ |Optional| **Variant**|Left position of the shape on the page.|
|_Top_ |Optional| **Variant**|Top position of the shape on the page.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **MoveToPage** method to move the first shape in the **Shapes** collection on the first page of a publication to the same relative location on the second page of the publication.

This code assumes that the current publication contains at least two pages, and that there is at least one shape on the first page of the publication.

```vb
Public Sub MoveToPage_Example() 
 
 Dim pubShape As Publisher.Shape 
 
 Set pubShape = ThisDocument.Pages(1).Shapes(1) 
 
 pubShape.MoveToPage 2 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]