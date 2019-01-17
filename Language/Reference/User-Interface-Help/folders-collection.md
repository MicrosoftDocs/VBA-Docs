---
title: Folders collection
keywords: vblr6.chm2181929
f1_keywords:
- vblr6.chm2181929
ms.prod: office
api_name:
- Office.Folders
ms.assetid: 84c95d58-9183-4820-bd45-817164497234
ms.date: 11/12/2018
localization_priority: Normal
---


# Folders collection

Collection of all **[Folder](folder-object.md)** objects contained within a **Folder** object.

## Remarks

The following code illustrates how to get a **Folders** collection and how to iterate the collection by using the **[For Each...Next](for-eachnext-statement.md)** statement.

```vb
Sub ShowFolderList(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.SubFolders
    For Each f1 in fc
        s = s & f1.name 
        s = s &  vbCrLf
    Next
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]