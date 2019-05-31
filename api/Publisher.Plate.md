---
title: Plate object (Publisher)
keywords: vbapb10.chm2949119
f1_keywords:
- vbapb10.chm2949119
ms.prod: publisher
api_name:
- Publisher.Plate
ms.assetid: f7d7dbb1-a6a4-780f-814e-8e95aaaeeeea
ms.date: 06/01/2019
localization_priority: Normal
---


# Plate object (Publisher)

Represents a single printer's plate. The **Plate** object is a member of the **[Plates](Publisher.Plates.md)** collection.
 
## Remarks

Use the **[Add](Publisher.Plates.Add.md)** method of the **Plates** collection to create a new plate.

Use the **[FindPlateByInkName](Publisher.Plates.FindPlateByInkName.md)** method to return a specific plate by referencing its ink name. Process colors are assigned different index numbers in the **Plates** collection than in the **[PrintablePlates](Publisher.PrintablePlates.md)** collection. 

Use the **FindPlateByInkName** method to ensure that the desired **Plate** object is accessed.

## Example

This example creates a new spot-color plate collection and adds a plate to it.

```vb
Sub AddNewPlates() 
 Dim plts As Plates 
 Set plts = ActiveDocument.CreatePlateCollection(Mode:=pbColorModeSpot) 
 plts.Add 
 With plts(1) 
 .Color.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .Luminance = 4 
 End With 
End Sub
```


## Methods

- [ConvertToProcess](Publisher.Plate.ConvertToProcess.md)
- [Delete](Publisher.Plate.Delete.md)

## Properties

- [Application](Publisher.Plate.Application.md)
- [Color](Publisher.Plate.Color.md)
- [Index](Publisher.Plate.Index.md)
- [InkName](Publisher.Plate.InkName.md)
- [InUse](Publisher.Plate.InUse.md)
- [Luminance](Publisher.Plate.Luminance.md)
- [Name](Publisher.Plate.Name.md)
- [Parent](Publisher.Plate.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]