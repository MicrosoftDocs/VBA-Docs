---
title: Plate Object (Publisher)
keywords: vbapb10.chm2949119
f1_keywords:
- vbapb10.chm2949119
ms.prod: publisher
api_name:
- Publisher.Plate
ms.assetid: f7d7dbb1-a6a4-780f-814e-8e95aaaeeeea
ms.date: 06/08/2017
---


# Plate Object (Publisher)

Represents a single printer's plate. The  **Plate** object is a member of the **[Plates](Publisher.Plates.md)** collection.
 


## Example

Use the  **[Add](Publisher.Plates.Add.md)** method of the **[Plates](Publisher.Plates.md)** collection to create a new plate. This example creates a new spot-color plate collection and adds a plate to it.
 

 

```
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

Use the  **[FindPlateByInkName](Publisher.Plates.FindPlateByInkName.md)** method to return a specific plate by referencing its ink name. Process colors are assigned different index numbers in the **Plates** collection than in the **[PrintablePlates](Publisher.PrintablePlates.md)** collection. Use the **FindPlateByInkName** method to insure the desired **Plate** or **[PrintablePlate](Publisher.PrintablePlate.md)** object is accessed.
 

 

## Methods



|**Name**|
|:-----|
|[ConvertToProcess](Publisher.Plate.ConvertToProcess.md)|
|[Delete](Publisher.Plate.Delete.md)|

## Properties



|**Name**|
|:-----|
|[Application](Publisher.Plate.Application.md)|
|[Color](Publisher.Plate.Color.md)|
|[Index](Publisher.Plate.Index.md)|
|[InkName](Publisher.Plate.InkName.md)|
|[InUse](Publisher.Plate.InUse.md)|
|[Luminance](Publisher.Plate.Luminance.md)|
|[Name](Publisher.Plate.Name.md)|
|[Parent](plate-parent-property-publisher.md)|

