---
title: AdvancedPrintOptions.NegativeImage property (Publisher)
keywords: vbapb10.chm7077893
f1_keywords:
- vbapb10.chm7077893
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.NegativeImage
ms.assetid: 32a524ce-da31-8dfa-3286-c5d9c74367ca
ms.date: 06/04/2019
localization_priority: Normal
---


# AdvancedPrintOptions.NegativeImage property (Publisher)

**True** to print a negative image of the specified publication. The default is **False**. Read/write **Boolean**.


## Syntax

_expression_.**NegativeImage**

_expression_ A variable that represents an **[AdvancedPrintOptions](Publisher.AdvancedPrintOptions.md)** object.


## Return value

Boolean


## Remarks

This property is only accessible if the active printer is a PostScript printer. Returns a run-time error if a non-PostScript printer is specified. Use the **[IsPostscriptPrinter](Publisher.AdvancedPrintOptions.IsPostscriptPrinter.md)** property to determine if the specified printer is a PostScript printer.

This property is saved as an application setting and applied to future instances of Microsoft Publisher.

This property corresponds to the **Negative image** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.


## Example

The following example determines if the active printer is a PostScript printer. If it is, the active publication is set to print as a horizontally and vertically mirrored negative image of itself.

```vb
Sub PrepToPrintToFilmOnImagesetter() 
 
With ActiveDocument.AdvancedPrintOptions 
 If .IsPostscriptPrinter = True Then 
 .HorizontalFlip = True 
 .VerticalFlip = True 
 .NegativeImage = True 
 End If 
End With 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]