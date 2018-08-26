---
title: Drives Property
keywords: vblr6.chm2182030
f1_keywords:
- vblr6.chm2182030
ms.prod: office
api_name:
- Office.Drives
ms.assetid: 23534228-121c-23df-6ea6-c4715f86e312
ms.date: 06/08/2017
---


# Drives Property



 **Description**
Returns a  **Drives** collection consisting of all **Drive** objects available on the local machine.
<<<<<<< HEAD
 **Syntax**
 _object_. **Drives**
The  _object_ is always a **FileSystemObject**.
 **Remarks**
=======

## Syntax

_object_. **Drives**
The  _object_ is always a **FileSystemObject**.

## Remarks

>>>>>>> master
Removable-media drives need not have media inserted for them to appear in the  **Drives** collection.
You can iterate the members of the  **Drives** collection using a **For Each...Next** construct as illustrated in the following code:



```vb
Sub ShowDriveList
    Dim fs, d, dc, s, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each d in dc
<<<<<<< HEAD
        s = s &; d.DriveLetter &; " - " 
=======
        s = s & d.DriveLetter & " - " 
>>>>>>> master
        If d.DriveType = 3 Then
            n = d.ShareName
        Else
            n = d.VolumeName
        End If
<<<<<<< HEAD
        s = s &; n &; vbCrLf
=======
        s = s & n & vbCrLf
>>>>>>> master
    Next
    MsgBox s
End Sub
```


