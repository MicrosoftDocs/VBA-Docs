---
title: Plate.InkName Property (Publisher)
keywords: vbapb10.chm2883603
f1_keywords:
- vbapb10.chm2883603
ms.prod: publisher
api_name:
- Publisher.Plate.InkName
ms.assetid: 248c1529-2706-5458-a13f-def479d16132
ms.date: 06/08/2017
---


# Plate.InkName Property (Publisher)

Returns a  **PbInkName** constant that represents the name of the ink to be printed using this plate. Read-only.


## Syntax

 _expression_. **InkName**

 _expression_ A variable that represents a  **Plate** object.


## Remarks

The  **InkName** property value can be one of the ** [PbInkName](./overview/Publisher.md)** constants declared in the Microsoft Publisher type library.

Use the  **FindPlateByInkName** method of the **[PrintablePlates](Publisher.PrintablePlates.md)** collection to return a specific plate by referencing its ink name.


## Example

The following example returns a list of the printable plates currently in the collection for the active publication. The example assumes that separations have been specified as the active publication's print mode.


```vb
Sub ListPrintablePlates() 
 Dim pplTemp As PrintablePlates 
 Dim pplLoop As PrintablePlate 
 
 
 Set pplTemp = ActiveDocument.AdvancedPrintOptions.PrintablePlates 
<<<<<<< HEAD
 Debug.Print "There are " &; pplTemp.Count &; " printable plates in this publication." 
 
 For Each pplLoop In pplTemp 
 With pplLoop 
 Debug.Print "Printable Plate Name: " &; .Name 
 Debug.Print "Index: " &; .Index 
 Debug.Print "Ink Name: " &; .InkName 
 Debug.Print "Plate Angle: " &; .Angle 
 Debug.Print "Plate Frequency: " &; .Frequency 
 Debug.Print "Print Plate?: " &; .PrintPlate 
=======
 Debug.Print "There are " & pplTemp.Count & " printable plates in this publication." 
 
 For Each pplLoop In pplTemp 
 With pplLoop 
 Debug.Print "Printable Plate Name: " & .Name 
 Debug.Print "Index: " & .Index 
 Debug.Print "Ink Name: " & .InkName 
 Debug.Print "Plate Angle: " & .Angle 
 Debug.Print "Plate Frequency: " & .Frequency 
 Debug.Print "Print Plate?: " & .PrintPlate 
>>>>>>> master
 End With 
 Next pplLoop 
End Sub
```


