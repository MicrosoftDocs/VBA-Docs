---
title: WebOptions.UseDefaultFolderSuffix method (Excel)
keywords: vbaxl10.chm662084
f1_keywords:
- vbaxl10.chm662084
ms.prod: excel
api_name:
- Excel.WebOptions.UseDefaultFolderSuffix
ms.assetid: dbaf5fa4-449a-b549-d2a0-82f65497f6c0
ms.date: 05/18/2019
localization_priority: Normal
---


# WebOptions.UseDefaultFolderSuffix method (Excel)

Sets the folder suffix for the specified document to the default suffix for the language support that you have selected or installed.


## Syntax

_expression_.**UseDefaultFolderSuffix**

_expression_ A variable that represents a **[WebOptions](Excel.WebOptions.md)** object.


## Remarks

Microsoft Excel uses the folder suffix when you save a document as a webpage, use long file names, and choose to save supporting files in a separate folder (that is, if the **[UseLongFileNames](Excel.WebOptions.UseLongFileNames.md)** and **[OrganizeInFolder](Excel.WebOptions.OrganizeInFolder.md)** properties are set to **True**).

The suffix appears in the folder name after the document name. For example, if the document is called Book1 and the language is English, the folder name is Book1_files. The available folder suffixes are listed in the **[FolderSuffix](Excel.WebOptions.FolderSuffix.md)** property topic.


## Example

This example sets the folder suffix for the first workbook to the default suffix.

```vb
Workbooks(1).WebOptions.UseDefaultFolderSuffix
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]