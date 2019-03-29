---
title: PageSetup object (Excel)
keywords: vbaxl10.chm472072
f1_keywords:
- vbaxl10.chm472072
ms.prod: excel
api_name:
- Excel.PageSetup
ms.assetid: 2fd22df9-5987-f723-04a9-9a3f2e84ac81
ms.date: 03/30/2019
localization_priority: Priority
---


# PageSetup object (Excel)

Represents the page setup description.


## Remarks

The **PageSetup** object contains all page setup attributes (left margin, bottom margin, paper size, and so on) as properties.


## Example

Use the **[PageSetup](Excel.Worksheet.PageSetup.md)** property of the **Worksheet** object to return a **PageSetup** object. 

The following example sets the orientation to landscape mode and then prints the worksheet.

```vb
With Worksheets("Sheet1") 
 .PageSetup.Orientation = xlLandscape 
 .PrintOut 
End With
```

<br/>

The **With** statement makes it easier and faster to set several properties at the same time. The following example sets all the margins for worksheet one.

```vb
With Worksheets(1).PageSetup 
 .LeftMargin = Application.InchesToPoints(0.5) 
 .RightMargin = Application.InchesToPoints(0.75) 
 .TopMargin = Application.InchesToPoints(1.5) 
 .BottomMargin = Application.InchesToPoints(1) 
 .HeaderMargin = Application.InchesToPoints(0.5) 
 .FooterMargin = Application.InchesToPoints(0.5) 
End With
```


## Properties

- [AlignMarginsHeaderFooter](Excel.PageSetup.AlignMarginsHeaderFooter.md)
- [Application](Excel.PageSetup.Application.md)
- [BlackAndWhite](Excel.PageSetup.BlackAndWhite.md)
- [BottomMargin](Excel.PageSetup.BottomMargin.md)
- [CenterFooter](Excel.PageSetup.CenterFooter.md)
- [CenterFooterPicture](Excel.PageSetup.CenterFooterPicture.md)
- [CenterHeader](Excel.PageSetup.CenterHeader.md)
- [CenterHeaderPicture](Excel.PageSetup.CenterHeaderPicture.md)
- [CenterHorizontally](Excel.PageSetup.CenterHorizontally.md)
- [CenterVertically](Excel.PageSetup.CenterVertically.md)
- [Creator](Excel.PageSetup.Creator.md)
- [DifferentFirstPageHeaderFooter](Excel.PageSetup.DifferentFirstPageHeaderFooter.md)
- [Draft](Excel.PageSetup.Draft.md)
- [EvenPage](Excel.PageSetup.EvenPage.md)
- [FirstPage](Excel.PageSetup.FirstPage.md)
- [FirstPageNumber](Excel.PageSetup.FirstPageNumber.md)
- [FitToPagesTall](Excel.PageSetup.FitToPagesTall.md)
- [FitToPagesWide](Excel.PageSetup.FitToPagesWide.md)
- [FooterMargin](Excel.PageSetup.FooterMargin.md)
- [HeaderMargin](Excel.PageSetup.HeaderMargin.md)
- [LeftFooter](Excel.PageSetup.LeftFooter.md)
- [LeftFooterPicture](Excel.PageSetup.LeftFooterPicture.md)
- [LeftHeader](Excel.PageSetup.LeftHeader.md)
- [LeftHeaderPicture](Excel.PageSetup.LeftHeaderPicture.md)
- [LeftMargin](Excel.PageSetup.LeftMargin.md)
- [OddAndEvenPagesHeaderFooter](Excel.PageSetup.OddAndEvenPagesHeaderFooter.md)
- [Order](Excel.PageSetup.Order.md)
- [Orientation](Excel.PageSetup.Orientation.md)
- [Pages](Excel.PageSetup.Pages.md)
- [PaperSize](Excel.PageSetup.PaperSize.md)
- [Parent](Excel.PageSetup.Parent.md)
- [PrintArea](Excel.PageSetup.PrintArea.md)
- [PrintComments](Excel.PageSetup.PrintComments.md)
- [PrintErrors](Excel.PageSetup.PrintErrors.md)
- [PrintGridlines](Excel.PageSetup.PrintGridlines.md)
- [PrintHeadings](Excel.PageSetup.PrintHeadings.md)
- [PrintNotes](Excel.PageSetup.PrintNotes.md)
- [PrintQuality](Excel.PageSetup.PrintQuality.md)
- [PrintTitleColumns](Excel.PageSetup.PrintTitleColumns.md)
- [PrintTitleRows](Excel.PageSetup.PrintTitleRows.md)
- [RightFooter](Excel.PageSetup.RightFooter.md)
- [RightFooterPicture](Excel.PageSetup.RightFooterPicture.md)
- [RightHeader](Excel.PageSetup.RightHeader.md)
- [RightHeaderPicture](Excel.PageSetup.RightHeaderPicture.md)
- [RightMargin](Excel.PageSetup.RightMargin.md)
- [ScaleWithDocHeaderFooter](Excel.PageSetup.ScaleWithDocHeaderFooter.md)
- [TopMargin](Excel.PageSetup.TopMargin.md)
- [Zoom](Excel.PageSetup.Zoom.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]