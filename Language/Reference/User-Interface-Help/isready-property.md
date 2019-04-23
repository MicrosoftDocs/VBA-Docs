---
title: IsReady property (Visual Basic for Applications)
keywords: vblr6.chm2181959
f1_keywords:
- vblr6.chm2181959
ms.prod: office
api_name:
- Office.IsReady
ms.assetid: e4c0771b-ea30-1431-2106-ca53a13543f2
ms.date: 12/19/2018
localization_priority: Normal
---


# IsReady property

Returns **True** if the specified drive is ready; **False** if it is not.

## Syntax

object.**IsReady**

The object is always a **[Drive](drive-object.md)** object.

## Remarks

For removable-media drives and CD-ROM drives, **IsReady** returns **True** only when the appropriate media is inserted and ready for access.

The following code illustrates the use of the **IsReady** property.

```vb
Sub ShowDriveInfo(drvpath)
    Dim fs, d, s, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    Select Case d.DriveType
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    s = "Drive " & d.DriveLetter & ": - " & t
    If d.IsReady Then 
        s = s & vbCrLf & "Drive is Ready."
    Else
        s = s & vbCrLf & "Drive is not Ready."
    End If
    MsgBox s
End Sub
```


## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]