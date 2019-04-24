---
title: DefaultWebOptions object (Excel)
keywords: vbaxl10.chm659072
f1_keywords:
- vbaxl10.chm659072
ms.prod: excel
api_name:
- Excel.DefaultWebOptions
ms.assetid: 5bd1d870-e8d9-cac1-d7a7-3aeaf7c4c3cd
ms.date: 03/29/2019
localization_priority: Normal
---


# DefaultWebOptions object (Excel)

Contains global application-level attributes used by Microsoft Excel when you save a document as a webpage or open a webpage. You can return or set attributes either at the application (global) level or at the workbook level.


## Remarks

Workbook-level attribute settings override application-level attribute settings. Workbook-level attributes are contained in the **[WebOptions](Excel.WebOptions.md)** object.

> [!NOTE] 
> Attribute values can be different from one workbook to another, depending on the attribute value at the time the workbook was saved.


## Example

Use the **[DefaultWebOptions](Excel.Application.DefaultWebOptions.md)** property of the **Application** object to return the **DefaultWebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and sets the _strImageFileType_ variable accordingly.

```vb
Set objAppWebOptions = Application.DefaultWebOptions 
With objAppWebOptions 
 If .AllowPNG = True Then 
 strImageFileType = "PNG" 
 Else 
 strImageFileType = "JPG" 
 End If 
End With
```


## Properties

- [AllowPNG](Excel.DefaultWebOptions.AllowPNG.md)
- [AlwaysSaveInDefaultEncoding](Excel.DefaultWebOptions.AlwaysSaveInDefaultEncoding.md)
- [Application](Excel.DefaultWebOptions.Application.md)
- [CheckIfOfficeIsHTMLEditor](Excel.DefaultWebOptions.CheckIfOfficeIsHTMLEditor.md)
- [Creator](Excel.DefaultWebOptions.Creator.md)
- [DownloadComponents](Excel.DefaultWebOptions.DownloadComponents.md)
- [Encoding](Excel.DefaultWebOptions.Encoding.md)
- [FolderSuffix](Excel.DefaultWebOptions.FolderSuffix.md)
- [Fonts](Excel.DefaultWebOptions.Fonts.md)
- [LoadPictures](Excel.DefaultWebOptions.LoadPictures.md)
- [LocationOfComponents](Excel.DefaultWebOptions.LocationOfComponents.md)
- [OrganizeInFolder](Excel.DefaultWebOptions.OrganizeInFolder.md)
- [Parent](Excel.DefaultWebOptions.Parent.md)
- [PixelsPerInch](Excel.DefaultWebOptions.PixelsPerInch.md)
- [RelyOnCSS](Excel.DefaultWebOptions.RelyOnCSS.md)
- [RelyOnVML](Excel.DefaultWebOptions.RelyOnVML.md)
- [SaveHiddenData](Excel.DefaultWebOptions.SaveHiddenData.md)
- [SaveNewWebPagesAsWebArchives](Excel.DefaultWebOptions.SaveNewWebPagesAsWebArchives.md)
- [ScreenSize](Excel.DefaultWebOptions.ScreenSize.md)
- [TargetBrowser](Excel.DefaultWebOptions.TargetBrowser.md)
- [UpdateLinksOnSave](Excel.DefaultWebOptions.UpdateLinksOnSave.md)
- [UseLongFileNames](Excel.DefaultWebOptions.UseLongFileNames.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]