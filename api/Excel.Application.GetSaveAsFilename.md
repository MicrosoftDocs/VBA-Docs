---
title: Application.GetSaveAsFilename method (Excel)
keywords: vbaxl10.chm133143
f1_keywords:
- vbaxl10.chm133143
ms.prod: excel
api_name:
- Excel.Application.GetSaveAsFilename
ms.assetid: 2ad52070-22d7-a755-9267-daaa5edbbb0d
ms.date: 04/04/2019
localization_priority: Priority
---


# Application.GetSaveAsFilename method (Excel)

Displays the standard **Save As** dialog box and gets a file name from the user without actually saving any files.


## Syntax

_expression_.**GetSaveAsFilename** (_InitialFilename_, _FileFilter_, _FilterIndex_, _Title_, _ButtonText_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _InitialFilename_|Optional| **Variant**|Specifies the suggested file name. If this argument is omitted, Microsoft Excel uses the active workbook's name.|
| _FileFilter_|Optional| **Variant**|A string specifying file filtering criteria. Max length is 255 characters, otherwise the method returns Error 2015.|
| _FilterIndex_|Optional| **Variant**|Specifies the index number of the default file filtering criteria, from 1 to the number of filters specified in  _FileFilter_. If this argument is omitted or greater than the number of filters present, the first file filter is used.|
| _Title_|Optional| **Variant**|Specifies the title of the dialog box. If this argument is omitted, the default title is used.|
| _ButtonText_|Optional| **Variant**|Macintosh only.|

## Return value

Variant


## Remarks

This string passed in the _FileFilter_ argument consists of pairs of file filter strings followed by the MS-DOS wildcard file filter specification, with each part and each pair separated by commas. Each separate pair is listed in the **Files of type** drop-down list box. For example, the following string specifies two file filters&mdash;text and addin: 

`"Text Files (*.txt), *.txt, Add-In Files (*.xla), *.xla"`

To use multiple MS-DOS wildcard expressions for a single file filter type, separate the wildcard expressions with semicolons; for example, `"Visual Basic Files (*.bas; *.txt),*.bas;*.txt"`.

This method returns the selected file name or the name entered by the user. The returned name may include a path specification. Returns **False** if the user cancels the dialog box.

This method may change the current drive or folder.

When InitialFilename is used with an extension and a filter is applied, this extension must match the filter extension, otherwise the effective InitialFilename displayed in the dialog box will be an empty string.

## Example

This example displays the **Save As** dialog box, with the file filter set to text files. If the user chooses a file name, the example displays that file name in a message box.


```vb
fileSaveName = Application.GetSaveAsFilename( _ 
 fileFilter:="Text Files (*.txt), *.txt") 
If fileSaveName <> False Then 
 MsgBox "Save as " & fileSaveName 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
