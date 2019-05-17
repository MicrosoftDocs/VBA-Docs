---
title: WebOptions.OrganizeInFolder property (Excel)
keywords: vbaxl10.chm662074
f1_keywords:
- vbaxl10.chm662074
ms.prod: excel
api_name:
- Excel.WebOptions.OrganizeInFolder
ms.assetid: 9df9aff2-3a24-3e1f-db3e-7280b50b806b
ms.date: 05/18/2019
localization_priority: Normal
---


# WebOptions.OrganizeInFolder property (Excel)

**True** if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a webpage. **False** if supporting files are saved in the same folder as the webpage. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**OrganizeInFolder**

_expression_ A variable that represents a **[WebOptions](Excel.WebOptions.md)** object.


## Remarks

The new folder is created in the folder where you have saved the webpage, and is named after the document. If long file names are used, a suffix is added to the folder name. The **[FolderSuffix](Excel.WebOptions.FolderSuffix.md)** property returns the folder suffix for the language support that you have selected or installed, or the default folder suffix.

If you save a document that was previously saved with the **OrganizeInFolder** property set to a different value, Microsoft Excel automatically moves the supporting files into or out of the folder, as appropriate.

If you don't use long file names (that is, if the **[UseLongFileNames](Excel.WebOptions.UseLongFileNames.md)** property is set to **False**), Microsoft Excel automatically saves any supporting files in a separate folder. The files cannot be saved in the same folder as the webpage.


## Example

This example specifies that all supporting files are saved in the same folder when the document is saved as a webpage.

```vb
Application.DefaultWebOptions.OrganizeInFolder = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]