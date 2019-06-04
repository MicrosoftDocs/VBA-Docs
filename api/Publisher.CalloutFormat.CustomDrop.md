---
title: CalloutFormat.CustomDrop method (Publisher)
keywords: vbapb10.chm2490385
f1_keywords:
- vbapb10.chm2490385
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.CustomDrop
ms.assetid: 65fc7309-acd0-5bdd-6bb0-1b6c41968775
ms.date: 06/05/2019
localization_priority: Normal
---


# CalloutFormat.CustomDrop method (Publisher)

Sets the vertical distance from the edge of the text bounding box to the place where the callout line attaches to the text box.


## Syntax

_expression_.**CustomDrop** (_Drop_)

_expression_ A variable that represents a **[CalloutFormat](Publisher.CalloutFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Drop_|Required| **Variant**|The drop distance. Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Remarks

The drop distance is normally measured from the top of the text box. However, if the **[AutoAttach](Publisher.CalloutFormat.AutoAttach.md)** property is set to **True**, and the text box is to the left of the origin of the callout line (the place to which the callout points), the drop distance is measured from the bottom of the text box.


## Example

This example sets the custom drop distance to 14 points, and specifies that the drop distance always be measured from the top. For the example to work, the third shape in the active publication must be a callout.

```vb
With ActiveDocument.Pages(1).Shapes(3).Callout 
 .CustomDrop Drop:=14 
 .AutoAttach = False 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]