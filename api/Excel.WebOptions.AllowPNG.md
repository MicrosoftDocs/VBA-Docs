---
title: WebOptions.AllowPNG property (Excel)
keywords: vbaxl10.chm662078
f1_keywords:
- vbaxl10.chm662078
ms.prod: excel
api_name:
- Excel.WebOptions.AllowPNG
ms.assetid: 4fad6401-af54-ad7f-a46f-8110e8c00ad4
ms.date: 05/18/2019
localization_priority: Normal
---


# WebOptions.AllowPNG property (Excel)

**True** if Portable Network Graphics (PNG) is allowed as an image format when you save documents as a webpage. **False** if PNG is not allowed as an output format. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**AllowPNG**

_expression_ A variable that represents a **[WebOptions](Excel.WebOptions.md)** object.


## Remarks

If you save images in the PNG format as opposed to any other file format, you might improve the image quality or reduce the size of those image files, and therefore decrease the download time, assuming that the web browsers that you are targeting support the PNG format.


## Example

This example enables PNG as an output format for the first workbook.

```vb
Application.Workbooks(1).WebOptions.AllowPNG = True
```

<br/>

Alternatively, PNG can be enabled as the global default for the application for newly created documents.

```vb
Application.DefaultWebOptions.AllowPNG = True
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]