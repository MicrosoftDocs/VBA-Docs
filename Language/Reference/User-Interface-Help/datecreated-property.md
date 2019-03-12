---
title: DateCreated property (Visual Basic for Applications)
keywords: vblr6.chm2181973
f1_keywords:
- vblr6.chm2181973
ms.prod: office
api_name:
- Office.DateCreated
ms.assetid: 2b176d36-d598-f922-ceba-989411368253
ms.date: 12/19/2018
localization_priority: Normal
---


# DateCreated property

Returns the date and time that the specified file or folder was created. Read-only.

## Syntax

_object_.**DateCreated**

The _object_ is always a **[File](file-object.md)** or **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **DateCreated** property with a file.

```vb
Sub ShowFileInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = "Created: " & f.DateCreated
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
