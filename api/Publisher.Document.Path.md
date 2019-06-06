---
title: Document.Path property (Publisher)
keywords: vbapb10.chm196644
f1_keywords:
- vbapb10.chm196644
ms.prod: publisher
api_name:
- Publisher.Document.Path
ms.assetid: 01926d63-e59e-5aad-3cb9-143166d253a5
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.Path property (Publisher)

Returns a **String** indicating the full path to the file of the saved active publication, not including the last separator or file name.


## Syntax

_expression_.**Path**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Remarks

The **[FullName](Publisher.Document.FullName.md)** property can be used to return both the path and file name.


## Example

The following example demonstrates the differences between the **Path**, **Name**, and **FullName** properties. This example is best illustrated if the publication is saved in a folder other than the default.

```vb
Sub PathNames() 
 
 Dim strPath As String 
 Dim strName As String 
 Dim strFullName As String 
 
 strPath = Application.ActiveDocument.Path 
 strName = Application.ActiveDocument.Name 
 strFullName = Application.ActiveDocument.FullName 
 
 ' Note the file name & path differences 
 ' while executing. 
 MsgBox "The path is: " & strPath 
 MsgBox "The file name is: " & strName 
 MsgBox "The path & file name are: " & strFullName 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]