---
title: Drives property (Visual Basic for Applications)
keywords: vblr6.chm2182030
f1_keywords:
- vblr6.chm2182030
ms.prod: office
api_name:
- Office.Drives
ms.assetid: 23534228-121c-23df-6ea6-c4715f86e312
ms.date: 12/19/2018
localization_priority: Normal
---


# Drives property

Returns a **[Drives](drives-collection.md)** collection consisting of all **[Drive](drive-object.md)** objects available on the local machine.

## Syntax

_object_.**Drives**

The _object_ is always a **[FileSystemObject](filesystemobject-object.md)**.

## Remarks

Removable-media drives need not have media inserted for them to appear in the **Drives** collection.

You can iterate the members of the **Drives** collection by using a **[For Each...Next](for-eachnext-statement.md)** construct as illustrated in the following code.

```vb
Sub ShowDriveList
    Dim fs, d, dc, s, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each d in dc
        s = s & d.DriveLetter & " - " 
        If d.DriveType = 3 Then
            n = d.ShareName
        Else
            n = d.VolumeName
        End If
        s = s & n & vbCrLf
    Next
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]