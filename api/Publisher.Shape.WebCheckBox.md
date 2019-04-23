---
title: Shape.WebCheckBox property (Publisher)
keywords: vbapb10.chm2228344
f1_keywords:
- vbapb10.chm2228344
ms.prod: publisher
api_name:
- Publisher.Shape.WebCheckBox
ms.assetid: 13796525-584f-7109-5dea-1f2baf1efda7
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.WebCheckBox property (Publisher)

Returns the  **[WebCheckBox](Publisher.WebCheckBox.md)** object associated with the specified shape.


## Syntax

_expression_.**WebCheckBox**

 _expression_ A variable that represents a  **Shape** object.


## Return value

WebCheckBox


## Example

This example creates a new Web check box and specifies that its default state is checked.


```vb
Dim shpNew As Shape 
Dim wcbTemp As WebCheckBox 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=123, Width:=17, Height:=12) 
 
Set wcbTemp = shpNew.WebCheckBox 
 
wcbTemp.Selected = msoTrue
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]