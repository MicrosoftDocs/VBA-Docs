---
title: Application.FileValidation Property (Excel)
keywords: vbaxl10.chm133335
f1_keywords:
- vbaxl10.chm133335
ms.prod: excel
api_name:
- Excel.Application.FileValidation
ms.assetid: 6ec989d0-2ed8-b4d9-997c-4f91507e6fca
ms.date: 06/08/2017
---


# Application.FileValidation Property (Excel)

Returns or sets how Excel will validate files before opening them. Read/write


## Syntax

 _expression_. `FileValidation`

 _expression_ A variable that represents an '[Application](Excel.Application(object).md)' object.


### Return value

 **[MsoFileValidationMode](Office.MsoFileValidationMode.md)**


## Remarks

Files that do not pass validation will be opened in a  **Protected View** window. If you set the **FileValidation** property, that setting will remain in effect for the entire session the application is open.


## See also


[Application Object](Excel.Application(object).md)

