---
title: DefaultWebOptions.AllowPNG property (Excel)
keywords: vbaxl10.chm660082
f1_keywords:
- vbaxl10.chm660082
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.AllowPNG
ms.assetid: b4cdab42-25ed-e313-ebf2-fc9abd68474b
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.AllowPNG property (Excel)

**True** if PNG (Portable Network Graphics) is allowed as an image format when you save documents as a webpage. **False** if PNG is not allowed as an output format. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**AllowPNG**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Remarks

If you save images in the PNG format as opposed to any other file format, you might improve the image quality or reduce the size of those image files, and therefore decrease the download time, assuming that the web browsers that you are targeting support the PNG format.


## Example

Alternatively, PNG can be enabled as the global default for the application for newly created documents.

```vb
Application.DefaultWebOptions.AllowPNG = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]