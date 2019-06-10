---
title: Options.PathForPictures property (Publisher)
keywords: vbapb10.chm1048596
f1_keywords:
- vbapb10.chm1048596
ms.prod: publisher
api_name:
- Publisher.Options.PathForPictures
ms.assetid: e66c8c86-f049-0f32-0a0d-60fd37470708
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.PathForPictures property (Publisher)

Returns a **String** that represents the default path for picture files. Read.


## Syntax

_expression_.**PathForPictures**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

String


## Example

This example places the default path for picture files in a string and then uses the path string to add the specified file to the active publication. Note that `FileName` must be replaced with a valid file name for this example to work.

```vb
Sub InsertNewPicture() 
 Dim strPicPath As String 
 
 strPicPath = Options.PathForPictures 
 
 ActiveDocument.Pages(1).Shapes.AddPicture FileName:=strPicPath _ 
 & "FileName", LinktoFile:=msoFalse, _ 
 SaveWithDocument:=msoTrue, Left:=50, Top:=50, Height:=200 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]