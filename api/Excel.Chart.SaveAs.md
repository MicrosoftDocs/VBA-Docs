---
title: Chart.SaveAs method (Excel)
keywords: vbaxl10.chm149172
f1_keywords:
- vbaxl10.chm149172
api_name:
- Excel.Chart.SaveAs
ms.assetid: 21e2a786-1df2-21ea-f32f-81e07dc2261c
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Chart.SaveAs method (Excel)

Saves changes to the chart or worksheet in a different file.


## Syntax

_expression_.**SaveAs** (_FileName_, _FileFormat_, _Password_, _WriteResPassword_, _ReadOnlyRecommended_, _CreateBackup_, _AddToMru_, _TextCodepage_, _TextVisualLayout_, _Local_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**| **Variant**. A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.|
| _FileFormat_|Optional| **Variant**|The file format to use when you save the file. For a list of valid choices, see the **[FileFormat](Excel.Workbook.FileFormat.md)** property. For an existing file, the default format is the last file format specified; for a new file, the default is the format of the version of Excel being used.|
| _Password_|Optional| **Variant**|A case-sensitive string (no more than 15 characters) that indicates the protection password to be given to the file.|
| _WriteResPassword_|Optional| **Variant**|A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only.|
| _ReadOnlyRecommended_|Optional| **Variant**| **True** to display a message when the file is opened, recommending that the file be opened as read-only.|
| _CreateBackup_|Optional| **Variant**| **True** to create a backup file.|
| _AddToMru_|Optional| **Variant**| **True** to add this workbook to the list of recently used files. The default value is **False**.|
| _TextCodepage_|Optional| **Variant**|Not used in U.S. English Microsoft Excel.|
| _TextVisualLayout_|Optional| **Variant**|Not used in U.S. English Microsoft Excel.|
| _Local_|Optional| **Variant**| **True** saves files against the language of Microsoft Excel (including control panel settings). **False** (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically US English unless the VBA project where **Workbooks.Open** is run from is an old internationalized XL5/95 VBA project).|

## Remarks

Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: **Y6dh!et5**. Weak password: **House27**. Use a strong password that you can remember so that you don't have to write it down.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]