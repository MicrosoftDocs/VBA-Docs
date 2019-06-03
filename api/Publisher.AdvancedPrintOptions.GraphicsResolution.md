---
title: AdvancedPrintOptions.GraphicsResolution property (Publisher)
keywords: vbapb10.chm7077909
f1_keywords:
- vbapb10.chm7077909
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.GraphicsResolution
ms.assetid: 1e4e06aa-327b-5689-ff97-eea9f866260a
ms.date: 06/04/2019
localization_priority: Normal
---


# AdvancedPrintOptions.GraphicsResolution property (Publisher)

Returns or sets a **[PbPrintGraphics](publisher.pbprintgraphics.md)** constant representing the resolution at which the inserted graphics are to be printed in the specified publication. Read/write.


## Syntax

_expression_.**GraphicsResolution**

_expression_ A variable that represents an **[AdvancedPrintOptions](Publisher.AdvancedPrintOptions.md)** object.


## Return value

**PbPrintGraphics**


## Remarks

Setting this property only affects inserted pictures (whether linked or embedded), and clip art. Autoshapes and border art will always be printed.

Printing boxes in place of graphics is useful when printing a quick proof of the layout that only shows the positioning of pictures.

This property corresponds to the **Graphics** controls on the **Graphics and Fonts** tab of the **Advanced Print Settings** dialog box.

The **GraphicsResolution** property value can be one of the **PbPrintGraphics** constants declared in the Microsoft Publisher type library.


## Example

The following example sets the graphics to print as boxes in the active publication.

```vb
Sub PrintGraphicAsBoxes 
 With ActiveDocument.AdvancedPrintOptions 
 If .GraphicsResolution <> pbPrintNoGraphics Then 
 .GraphicsResolution = pbPrintNoGraphics 
 End If 
 End With 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]