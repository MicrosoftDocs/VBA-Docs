---
title: PrintableRect object (Publisher)
keywords: vbapb10.chm7602175
f1_keywords:
- vbapb10.chm7602175
ms.prod: publisher
api_name:
- Publisher.PrintableRect
ms.assetid: fd99e9d4-81d9-63ae-78ca-f7a16b031239
ms.date: 06/01/2019
localization_priority: Normal
---


# PrintableRect object (Publisher)

Represents the sheet area within which the specified printer will print. The printable rectangle is determined by the printer based on the sheet size specified. The printable rectangle of the printer sheet should not be confused with the area within the margins of the publication page; it may be larger or smaller than the publication page.
 

## Remarks

In cases in which the printer sheet and the publication page size are identical, the publication page is centered on the printer sheet and none of the printer's marks print, even if they are selected.
 
Use the **[Printer.PrintableRect](Publisher.Printer.PrintableRect.md)** property to return a **PrintableRect** object. 
 

## Example

The following example returns printable rectangle boundaries for the printer sheet of the active publication.

```vb
Sub ListPrintableRectBoundaries() 
 
With ActiveDocument.AdvancedPrintOptions.PrintableRect 
 
 Debug.Print "Printable area is " & _ 
 PointsToInches(.Width) & _ 
 " by " & PointsToInches(.Height) & " inches." 
 Debug.Print "Left Boundary: " & PointsToInches(.Left) & _ 
 " inches (from left)." 
 Debug.Print "Right Boundary: " & PointsToInches(.Left + .Width) & _ 
 " inches (from left)." 
 Debug.Print "Top Boundary: " & PointsToInches(.Top) & _ 
 " inches(from top)." 
 Debug.Print "Bottom Boundary: " & PointsToInches(.Top + .Height) & _ 
 " inches(from top)." 
 
End With 
 
End Sub 

```


## Properties

- [Application](Publisher.PrintableRect.Application.md)
- [Height](Publisher.PrintableRect.Height.md)
- [Left](Publisher.PrintableRect.Left.md)
- [Parent](Publisher.PrintableRect.Parent.md)
- [Top](Publisher.PrintableRect.Top.md)
- [Width](Publisher.PrintableRect.Width.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]