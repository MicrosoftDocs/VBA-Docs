---
title: AdvancedPrintOptions object (Publisher)
keywords: vbapb10.chm7143423
f1_keywords:
- vbapb10.chm7143423
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions
ms.assetid: 61f776cc-dc3e-61b6-057a-125ad15146c8
ms.date: 05/31/2019
localization_priority: Normal
---


# AdvancedPrintOptions object (Publisher)

Represents the advanced print settings for a publication.

## Remarks

The properties of the **AdvancedPrintOptions** object correspond to the options available on the tabs of the **Advanced Print Settings** dialog box.

Use the **[AdvancedPrintOptions](publisher.document.advancedprintoptions.md)** property of the **Document** object to return an **AdvancedPrintOptions** object.
 
## Example

The following example tests to determine if the active publication has been set to print as separations. If it has, it is set to print only plates for the inks actually used in the publication, and to not print plates for any pages where a color is not used.

```vb
Sub PrintOnlyInksUsed 
 With ActiveDocument.AdvancedPrintOptions 
 If .PrintMode = pbPrintModeSeparations Then 
 .InksToPrint = pbInksToPrintUsed 
 .PrintBlankPlates = False 
 End If 
 End With 
End Sub
```
<!--There is no PbInkName enumeration-->

## Properties

- [AllowBleeds](Publisher.AdvancedPrintOptions.AllowBleeds.md)
- [Application](Publisher.AdvancedPrintOptions.Application.md)
- [BackSideInsertFaceUp](Publisher.AdvancedPrintOptions.BackSideInsertFaceUp.md)
- [GraphicsResolution](Publisher.AdvancedPrintOptions.GraphicsResolution.md)
- [HorizontalFlip](Publisher.AdvancedPrintOptions.HorizontalFlip.md)
- [IsPostscriptPrinter](Publisher.AdvancedPrintOptions.IsPostscriptPrinter.md)
- [ManualFeedAlign](Publisher.AdvancedPrintOptions.ManualFeedAlign.md)
- [ManualFeedDirection](Publisher.AdvancedPrintOptions.ManualFeedDirection.md)
- [NegativeImage](Publisher.AdvancedPrintOptions.NegativeImage.md)
- [PageRotated](Publisher.AdvancedPrintOptions.PageRotated.md)
- [Parent](Publisher.AdvancedPrintOptions.Parent.md)
- [PrintBleedMarks](Publisher.AdvancedPrintOptions.PrintBleedMarks.md)
- [PrintCropMarks](Publisher.AdvancedPrintOptions.PrintCropMarks.md)
- [PrintDensityBars](Publisher.AdvancedPrintOptions.PrintDensityBars.md)
- [PrintJobInformation](Publisher.AdvancedPrintOptions.PrintJobInformation.md)
- [PrintRegistrationMarks](Publisher.AdvancedPrintOptions.PrintRegistrationMarks.md)
- [Resolution](Publisher.AdvancedPrintOptions.Resolution.md)
- [UseOnlyPublicationFonts](Publisher.AdvancedPrintOptions.UseOnlyPublicationFonts.md)
- [VerticalFlip](Publisher.AdvancedPrintOptions.VerticalFlip.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]