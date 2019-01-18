---
title: DefaultWebOptions object (Excel)
keywords: vbaxl10.chm659072
f1_keywords:
- vbaxl10.chm659072
ms.prod: excel
api_name:
- Excel.DefaultWebOptions
ms.assetid: 5bd1d870-e8d9-cac1-d7a7-3aeaf7c4c3cd
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions object (Excel)

Contains global application-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page. You can return or set attributes either at the application (global) level or at the workbook level.


## Remarks

 Workbook-level attribute settings override application-level attribute settings. Workbook-level attributes are contained in the **[WebOptions](Excel.WebOptions.md)** object.


 **Note**  Attribute values can be different from one workbook to another, depending on the attribute value at the time the workbook was saved.


## Example

Use the  **[DefaultWebOptions](Excel.Application.DefaultWebOptions.md)** property to return the **DefaultWebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and sets the _strImageFileType_ variable accordingly.


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


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]