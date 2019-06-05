---
title: Document.FullName property (Publisher)
keywords: vbapb10.chm196625
f1_keywords:
- vbapb10.chm196625
ms.prod: publisher
api_name:
- Publisher.Document.FullName
ms.assetid: 137e4310-8431-ed2a-503a-c225378a9a74
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.FullName property (Publisher)

Returns a **String** representing the full file name of the saved active publication, including its path and file name. Read-only.


## Syntax

_expression_.**FullName**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

String


## Remarks

The **FullName** property can be used to return both the path and file name as returned by the **[Path](Publisher.Document.Path.md)** and **[Name](Publisher.Document.Name.md)** properties.


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