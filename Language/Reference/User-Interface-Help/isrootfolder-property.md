---
title: IsRootFolder property (Visual Basic for Applications)
keywords: vblr6.chm2182069
f1_keywords:
- vblr6.chm2182069
ms.prod: office
api_name:
- Office.IsRootFolder
ms.assetid: 4d47b8c1-9ca0-a6d4-996d-584d55033cc1
ms.date: 12/19/2018
localization_priority: Normal
---


# IsRootFolder property

Returns **True** if the specified folder is the root folder; **False** if it is not.

## Syntax

_object_.**IsRootFolder**

The _object_ is always a **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **IsRootFolder** property.

```vb
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
Sub DisplayLevelDepth(pathspec)
    Dim f, n
    Set f = fs.GetFolder(pathspec)
    If f.IsRootFolder Then
        MsgBox "The specified folder is the root folder."
    Else
        Do Until f.IsRootFolder
            Set f = f.ParentFolder
            n = n + 1
        Loop
        MsgBox "The specified folder is nested " & n & " levels deep."
    End If
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]