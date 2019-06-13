---
title: Shape.WebCommandButton property (Publisher)
keywords: vbapb10.chm2228340
f1_keywords:
- vbapb10.chm2228340
ms.prod: publisher
api_name:
- Publisher.Shape.WebCommandButton
ms.assetid: c20b937b-6f53-fdc1-830a-4044831c351a
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.WebCommandButton property (Publisher)

Returns the **[WebCommandButton](Publisher.WebCommandButton.md)** object associated with the specified shape.


## Syntax

_expression_.**WebCommandButton**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Return value

WebCommandButton


## Example

This example creates a web form **Submit** command button and sets the script path and file name to run when a user chooses the button.

```vb
Dim shpNew As Shape 
Dim wcbTemp As WebCommandButton 
 
Set shpNew = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36) 
 
Set wcbTemp = shpNew.WebCommandButton 
 
With wcbTemp 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "https://www.tailspintoys.com/" _ 
 & "scripts/ispscript.cgi" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]