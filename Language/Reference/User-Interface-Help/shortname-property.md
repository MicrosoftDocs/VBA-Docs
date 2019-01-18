---
title: ShortName property (Visual Basic for Applications)
keywords: vblr6.chm2181997
f1_keywords:
- vblr6.chm2181997
ms.prod: office
api_name:
- Office.ShortName
ms.assetid: 62d95787-61c7-777d-56d0-d17d4d8e0f18
ms.date: 12/19/2018
localization_priority: Normal
---


# ShortName property

Returns the short name used by programs that require the earlier 8.3 naming convention.

## Syntax

_object_.**ShortName**

The _object_ is always a **[File](file-object.md)** or **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **ShortName** property with a **File** object.

```vb
Sub ShowShortName(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = "The short name for " & "" & UCase(f.Name)
    s = s & "" & vbCrLf
    s = s & "is: " & "" & f.ShortName & ""
    MsgBox s, 0, "Short Name Info"
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]