---
title: Application.PathSeparator Property (Publisher)
keywords: vbapb10.chm131104
f1_keywords:
- vbapb10.chm131104
ms.prod: publisher
api_name:
- Publisher.Application.PathSeparator
ms.assetid: f8c07ce4-d171-9c5b-60ac-d544bf65e620
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PathSeparator Property (Publisher)

Returns a  **String** that represents the character used to separate folder names. Read-only.


## Syntax

 _expression_. **PathSeparator**

 _expression_ A variable that represents a  **Application** object.


## Return value

String


## Remarks

You can use  **PathSeparator** to build Web addresses even though they contain forward slashes (/).

The  **[FullName](Publisher.Document.FullName.md)** property returns the path and file name as a single string.

For worldwide compatibility, we recommend that you use this property when building paths, rather than referring explicitly to path separator characters in code (for example, "/").


## Example

This example displays the path and file name of the active document.


```vb
Sub PathFileName() 
 
 With Application 
 MsgBox "The name of the active document: " & vbLf & _ 
 .Path & .PathSeparator & ActiveDocument.Name 
 End With 
 
End Sub
```


## See also


 [Application Object](Publisher.Application.md)

