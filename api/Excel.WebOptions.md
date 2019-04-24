---
title: WebOptions object (Excel)
keywords: vbaxl10.chm661072
f1_keywords:
- vbaxl10.chm661072
ms.prod: excel
api_name:
- Excel.WebOptions
ms.assetid: d573637f-1891-4602-c961-091795e47356
ms.date: 04/03/2019
localization_priority: Normal
---


# WebOptions object (Excel)

Contains workbook-level attributes used by Microsoft Excel when you save a document as a webpage or open a webpage.


## Remarks

You can return or set attributes either at the application (global) level or at the workbook level. (Note that attribute values can be different from one workbook to another, depending on the attribute value at the time the workbook was saved.) Workbook-level attribute settings override application-level attribute settings. Application-level attributes are contained in the **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Example

Use the **[WebOptions](Excel.Workbook.WebOptions.md)** property of the **Workbook** object to return the **WebOptions** object. 

The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format, and then sets the `strImageFileType` variable accordingly.

```vb
Set objAppWebOptions = Workbooks(1).WebOptions 
With objAppWebOptions 
 If .AllowPNG = True Then 
 strImageFileType = "PNG" 
 Else 
 strImageFileType = "JPG" 
 End If 
End With
```

## Methods

- [UseDefaultFolderSuffix](Excel.WebOptions.UseDefaultFolderSuffix.md)

## Properties

- [AllowPNG](Excel.WebOptions.AllowPNG.md)
- [Application](Excel.WebOptions.Application.md)
- [Creator](Excel.WebOptions.Creator.md)
- [DownloadComponents](Excel.WebOptions.DownloadComponents.md)
- [Encoding](Excel.WebOptions.Encoding.md)
- [FolderSuffix](Excel.WebOptions.FolderSuffix.md)
- [LocationOfComponents](Excel.WebOptions.LocationOfComponents.md)
- [OrganizeInFolder](Excel.WebOptions.OrganizeInFolder.md)
- [Parent](Excel.WebOptions.Parent.md)
- [PixelsPerInch](Excel.WebOptions.PixelsPerInch.md)
- [RelyOnCSS](Excel.WebOptions.RelyOnCSS.md)
- [RelyOnVML](Excel.WebOptions.RelyOnVML.md)
- [ScreenSize](Excel.WebOptions.ScreenSize.md)
- [TargetBrowser](Excel.WebOptions.TargetBrowser.md)
- [UseLongFileNames](Excel.WebOptions.UseLongFileNames.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]