---
title: AtEndOfLine property (Visual Basic for Applications)
keywords: vblr6.chm2182071
f1_keywords:
- vblr6.chm2182071
ms.prod: office
api_name:
- Office.AtEndOfLine
ms.assetid: a5b02fc7-362c-474d-7238-64c0783277ce
ms.date: 12/19/2018
---


# AtEndOfLine Property



 **Description**
Read-only property that returns  **True** if the file pointer immediately precedes the end-of-line marker in a **[TextStream](textstream-object.md)** file; **False** if it does not.

## Syntax

_object_. **AtEndOfLine**
The  _object_ is always the name of a **[TextStream](textstream-object.md)** object.

## Remarks

The  **AtEndOfLine** property applies only to **[TextStream](textstream-object.md)** files that are open for reading; otherwise, an error occurs.
The following code illustrates the use of the  **AtEndOfLine** property:



```vb
Dim fs, a, retstring
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile("c:\testfile.txt", ForReading, False)
Do While a. AtEndOfLine <> True
    retstring = a.Read(1)
    ...
Loop
a.Close
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)
