---
title: Plates Object (Publisher)
keywords: vbapb10.chm2883583
f1_keywords:
- vbapb10.chm2883583
ms.prod: publisher
api_name:
- Publisher.Plates
ms.assetid: 7da44b06-c94f-dadc-da91-09b757d5a076
ms.date: 06/08/2017
localization_priority: Normal
---


# Plates Object (Publisher)

A collection of  **Plate** objects in a publication.
 


## Example

The  **Plates** collection is made up of **Plate** objects for the various publication color modes. Each publication can only use one color mode. For example, you can't specify the spot-color mode in a procedure and then later specify the process-color mode. Use the **[CreatePlateCollection](./overview/Publisher.md)** method of the **[Document](Publisher.Document.md)** object to specify which color mode to use in a publication's plate collection. Use the **[Add](Publisher.Plates.Add.md)** method of the **Plates** collection to add a new plate to the **Plates** collection. This example creates a new spot-color plate collection and adds a plate to it.
 

 

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

Use the  **[EnterColorMode](./overview/Publisher.md)** method of the **[Document](Publisher.Document.md)** object to the specify the color mode and the **Plates** collection to use with the color mode. Use the **[ColorMode](./overview/Publisher.md)** property to determine which color mode is in use in a publication. This example creates a spot-color plate collection, adds two plates to it, and then enters those plates into the spot-color mode.
 

 



```vb
Sub CreateSpotColorMode() 
 Dim plArray As Plates 
 
 With ThisDocument 
 'Creates a color plate collection, 
 'which contains one black plate by default 
 Set plArray = .CreatePlateCollection(Mode:=pbColorModeSpot) 
 
 'Sets the plate color to red 
 plArray(1).Color.RGB = RGB(255, 0, 0) 
 
 'Adds another plate, black by default and 
 'sets the plate color to green 
 plArray.Add 
 plArray(2).Color.RGB = RGB(0, 255, 0) 
 
 'Enters spot-color mode with above 
 'two plates in the plates array 
 .EnterColorMode Mode:=pbColorModeSpot, Plates:=plArray 
 End With 
End Sub
```

Use the  **[FindPlateByInkName](Publisher.Plates.FindPlateByInkName.md)** method to return a specific plate by referencing its ink name. Process colors are assigned different index numbers in the **Plates** collection than in the **[PrintablePlates](Publisher.PrintablePlates.md)** collection. Use the **FindPlateByInkName** method to insure the desired **[Plate](Publisher.Plate.md)** or **[PrintablePlate](Publisher.PrintablePlate.md)** object is accessed.
 

 

## Methods



|Name|
|:-----|
|[Add](Publisher.Plates.Add.md)|
|[FindPlateByInkName](Publisher.Plates.FindPlateByInkName.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.Plates.Application.md)|
|[Count](Publisher.Plates.Count.md)|
|[Item](Publisher.Plates.Item.md)|
|[Parent](Publisher.Plates.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]