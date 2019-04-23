---
title: Printer.RowSpacing property (Access)
keywords: vbaac10.chm12873
f1_keywords:
- vbaac10.chm12873
ms.prod: access
api_name:
- Access.Printer.RowSpacing
ms.assetid: 78d6a87d-53ae-9c35-3ca6-3b66cb162ecf
ms.date: 03/23/2019
localization_priority: Normal
---


# Printer.RowSpacing property (Access)

Returns or sets a **Long** indicating the horizontal space between detail sections in [twips](../language/glossary/vbe-glossary.md#twip). Read/write.


## Syntax

_expression_.**RowSpacing**

_expression_ A variable that represents a **[Printer](Access.Printer.md)** object.


## Example

The following example sets a variety of printer settings for the form specified in the _strFormname_ argument of the procedure.

```vb
Sub SetPrinter(strFormname As String) 
 
 DoCmd.OpenForm FormName:=strFormname, view:=acDesign, _ 
 datamode:=acFormEdit, windowmode:=acHidden 
 
 With Forms(form1).Printer 
 
 .TopMargin = 1440 
 .BottomMargin = 1440 
 .LeftMargin = 1440 
 .RightMargin = 1440 
 
 .ColumnSpacing = 360 
  .RowSpacing = 360 
 
 .ColorMode = acPRCMColor 
 .DataOnly = False 
 .DefaultSize = False 
 .ItemSizeHeight = 2880 
 .ItemSizeWidth = 2880 
 .ItemLayout = acPRVerticalColumnLayout 
 .ItemsAcross = 6 
 
 .Copies = 1 
 .Orientation = acPRORLandscape 
 .Duplex = acPRDPVertical 
 .PaperBin = acPRBNAuto 
 .PaperSize = acPRPSLetter 
 .PrintQuality = acPRPQMedium 
 
 End With 
 
 DoCmd.Close objecttype:=acForm, objectname:=strFormname, _ 
 Save:=acSaveYes 
 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]