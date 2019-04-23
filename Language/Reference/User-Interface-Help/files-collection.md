---
title: Files collection
keywords: vblr6.chm2181926
f1_keywords:
- vblr6.chm2181926
ms.prod: office
api_name:
- Office.Files
ms.assetid: 1c69f6df-debc-448a-6f22-a2a41d069dc4
ms.date: 11/12/2018
localization_priority: Normal
---


# Files collection

Collection of all **[File](file-object.md)** objects within a folder.

## Remarks

The following code illustrates how to get a **Files** collection and iterate the collection by using the **[For Each...Next](for-eachnext-statement.md)** statement.

```vb
Sub ShowFolderList(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 in fc
        s = s & f1.name 
        s = s & vbCrLf
    Next
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
