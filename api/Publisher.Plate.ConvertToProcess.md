---
title: Plate.ConvertToProcess method (Publisher)
keywords: vbapb10.chm2883601
f1_keywords:
- vbapb10.chm2883601
ms.prod: publisher
api_name:
- Publisher.Plate.ConvertToProcess
ms.assetid: 26476701-aa82-ca44-20c8-55a332a6539a
ms.date: 06/13/2019
localization_priority: Normal
---


# Plate.ConvertToProcess method (Publisher)

Converts the specified plate from spot color to process.


## Syntax

_expression_.**ConvertToProcess**

_expression_ A variable that represents a **[Plate](Publisher.Plate.md)** object.


## Remarks

The **ConvertToProcess** method is only accessible if the publication's color mode has been set to process and spot color inks. 

Returns "Permission Denied" when applied to a spot color plate. When the color mode includes process color, the process color inks (black, magenta, yellow, and cyan) are the first four plates in the **[Plates](Publisher.Plates.md)** collection.

When a plate is converted from spot to process color, all colors in the publication based on the ink that the converted plate represents are converted to process colors.


## Example

The following example converts the specified spot color plate to process color. The example assumes that the publication's color mode has been specified as spot and process color, and that at least six plates have been defined for the publication.

```vb
Sub ChangePlateToProcess() 
 
 With ActiveDocument.Plates.Item(6) 
 .ConvertToProcess 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]