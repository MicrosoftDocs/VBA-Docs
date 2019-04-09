---
title: FileConverter.OpenFormat property (Word)
keywords: vbawd10.chm161021955
f1_keywords:
- vbawd10.chm161021955
ms.prod: word
api_name:
- Word.FileConverter.OpenFormat
ms.assetid: d5a83e1f-bbf6-d0f5-8223-c2140850bc27
ms.date: 06/08/2017
localization_priority: Normal
---


# FileConverter.OpenFormat property (Word)

Returns the file format of the specified file converter. Read-only  **Long**.


## Syntax

_expression_. `OpenFormat`

_expression_ Required. A variable that represents a '[FileConverter](Word.FileConverter.md)' object.


## Remarks

This property can be any valid  **WdOpenFormat** constant, or it can be a unique number that represents an external file converter.


## Example

This example displays the unique format value and the format name for the converters you can use to open documents.


```vb
For Each fc In FileConverters 
 If fc.CanOpen = True Then _ 
 MsgBox fc.OpenFormat & vbCr & fc.FormatName 
Next fc
```

This example opens the file named "Data.wp" by using the WordPerfect 6x file converter.




```vb
Documents.Open FileName:="C:\Data.wp", _ 
 Format:=FileConverters("WordPerfect6x").OpenFormat
```


## See also


[FileConverter Object](Word.FileConverter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]