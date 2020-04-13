---
title: FileConverter.CanOpen property (PowerPoint)
keywords: vbapp10.chm680002
f1_keywords:
- vbapp10.chm680002
ms.prod: powerpoint
api_name:
- PowerPoint.FileConverter.CanOpen
ms.assetid: 9a5a2fea-0f09-9dfe-c75a-e8811d53c27f
ms.date: 06/08/2017
localization_priority: Normal
---


# FileConverter.CanOpen property (PowerPoint)

 **True** if the specified file converter is designed to open files. Read-only **Boolean**.


## Syntax

_expression_. `CanOpen`

_expression_ A variable that represents a '[FileConverter](PowerPoint.FileConverter.md)' object.


## Remarks

The **[CanSave](PowerPoint.FileConverter.CanSave.md)** property returns **True** if the specified file converter can be used to save (export) files.


## Example

This example determines whether the first file converter is able to open files.


```vb
If FileConverters(1).CanOpen = True Then

    MsgBox FileConverters(1).FormatName & " can open files"

End If
```




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

This example determines whether the WordPerfect6x file converter can be used to open files. If the CanOpen property returns True, a document named "Test.wp" is opened.




```vb
If FileConverters("WordPerfect6x").CanOpen = True Then
    Documents.Open FileName:="C:\Test.wp", _
        Format:=FileConverters("WordPerfect6x").OpenFormat
End If
```


## See also


[FileConverter Object](PowerPoint.FileConverter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]