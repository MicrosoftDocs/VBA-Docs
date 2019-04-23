---
title: Files property (Visual Basic for Applications)
keywords: vblr6.chm2182095
f1_keywords:
- vblr6.chm2182095
ms.prod: office
api_name:
- Office.Files
ms.assetid: 80ee842f-759f-a018-c69c-4233d9714938
ms.date: 12/19/2018
localization_priority: Normal
---


# Files property

Returns a **Files** collection consisting of all **[File](file-object.md)** objects contained in the specified folder, including those with hidden and system file attributes set.

## Syntax

_object_.**Files**

The _object_ is always a **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **Files** property.

```vb
Sub ShowFileList(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 in fc
        s = s & f1.name 
        s = s &  vbCrLf
    Next
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]