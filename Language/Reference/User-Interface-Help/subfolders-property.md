---
title: SubFolders property (Visual Basic for Applications)
keywords: vblr6.chm2182070
f1_keywords:
- vblr6.chm2182070
ms.prod: office
api_name:
- Office.SubFolders
ms.assetid: 60bc795f-22f9-6846-00d3-05229f062099
ms.date: 12/19/2018
localization_priority: Normal
---


# SubFolders property

Returns a **[Folders](folders-collection.md)** collection consisting of all folders contained in a specified folder, including those with Hidden and System file attributes set.

## Syntax

_object_.**SubFolders**

The _object_ is always a **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **SubFolders** property.

```vb
Sub ShowFolderList(folderspec)
    Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
    For Each f1 in sf
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
