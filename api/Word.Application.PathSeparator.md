---
title: Application.PathSeparator property (Word)
keywords: vbawd10.chm158335072
f1_keywords:
- vbawd10.chm158335072
ms.prod: word
api_name:
- Word.Application.PathSeparator
ms.assetid: 29347a13-8edb-0b02-32c3-d091eb52c9f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PathSeparator property (Word)

Returns the character used to separate folder names. This property returns a backslash (\). Read-only  **String**.


## Syntax

_expression_. `PathSeparator`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


## Remarks

You can use  **PathSeparator** property to build web addresses even though they contain forward slashes (/).


> [!NOTE] 
> The  **[FullName](Word.Document.FullName.md)** property returns the path and file name, including the path separator, as a single string.


## Example

This example displays the path and file name of the active document.


```vb
MsgBox ActiveDocument.Path & Application.PathSeparator & _ 
 ActiveDocument.Name
```

If the first add-in is a template, this example unloads the template and opens it.




```vb
If Addins(1).Compiled = False Then 
 Addins(1).Installed = False 
 Documents.Open FileName:=AddIns(1).Path _ 
 & Application.PathSeparator _ 
 & AddIns(1).Name 
End If
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]