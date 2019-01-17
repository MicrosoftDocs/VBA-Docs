---
title: OpenTextFile method (Visual Basic for Applications)
keywords: vblr6.chm2182061
f1_keywords:
- vblr6.chm2182061
ms.prod: office
api_name:
- Office.OpenTextFile
ms.assetid: f44f7bc5-e48b-05f2-eb22-5b02701d449e
ms.date: 12/14/2018
localization_priority: Priority
---


# OpenTextFile method

Opens a specified file and returns a **[TextStream](textstream-object.md)** object that can be used to read from, write to, or append to the file.

## Syntax

_object_.**OpenTextFile** (_filename_, [ _iomode_, [ _create_, [ _format_ ]]])

<br/>

The **OpenTextFile** method has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _filename_|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) that identifies the file to open.|
| _iomode_|Optional. Indicates input/output mode. Can be one of three constants: **ForReading**, **ForWriting**, or **ForAppending**.|
| _create_|Optional. **Boolean** value that indicates whether a new file can be created if the specified _filename_ doesn't exist. The value is **True** if a new file is created; **False** if it isn't created. The default is **False**.|
| _format_|Optional. One of three **Tristate** values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.|

## Settings

The _iomode_ argument can have any of the following settings:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**ForReading**|1|Open a file for reading only. You can't write to this file.|
|**ForWriting**|2|Open a file for writing only. Use this mode to replace an existing file with new data. You can't read from this file.|
|**ForAppending**|8|Open a file and write to the end of the file. You can't read from this file.|

<br/>

The _format_ argument can have any of the following settings:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**TristateUseDefault**|-2|Opens the file by using the system default.|
|**TristateTrue**|-1|Opens the file as Unicode.|
|**TristateFalse**| 0|Opens the file as ASCII.|

## Remarks

The following code illustrates the use of the **OpenTextFile** method to open a file for appending text:

```vb
Sub OpenTextFileTest
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile("c:\testfile.txt", ForAppending, TristateFalse)
    f.Write "Hello world!"
    f.Close
End Sub
```


## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
