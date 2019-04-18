---
title: TextStream object
keywords: vblr6.chm2181930
f1_keywords:
- vblr6.chm2181930
ms.prod: office
api_name:
- Office.TextStream
ms.assetid: b1b78d3a-78b3-aee5-2efc-1e208e0858ac
ms.date: 11/12/2018
localization_priority: Normal
---


# TextStream object

Facilitates sequential access to file.

## Syntax

**TextStream**. { _property_ | _method_ }

The _property_ and _method_ arguments can be any of the properties and methods associated with the **TextStream** object. Note that in actual usage, **TextStream** is replaced by a variable placeholder representing the **TextStream** object returned from the **FileSystemObject**.

## Remarks

In the following code, `a` is the **TextStream** object returned by the **CreateTextFile** method on the **FileSystemObject**; **WriteLine** and **Close** are two methods of the **TextStream** object.

```vb
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("c:\testfile.txt", True)
a.WriteLine("This is a test.")
a.Close

```

## Methods

|Method|Description|
|:-----|:----------|
|[Close](close-method-textstream-object.md)|Closes an open TextStream file. |
|[Read](read-method.md)|Reads a specified number of characters from a TextStream file and returns the result. |
|[ReadAll](readall-method.md)|Reads an entire TextStream file and returns the result. |
|[ReadLine](readline-method.md)|Reads one line from a TextStream file and returns the result. |
|[Skip](skip-method.md)|Skips a specified number of characters when reading a TextStream file. |
|[SkipLine](skipline-method.md)|Skips the next line when reading a TextStream file. |
|[Write](write-method.md)|Writes a specified text to a TextStream file. |
|[WriteBlankLines](writeblanklines-method.md)|Writes a specified number of new-line characters to a TextStream file. |
|[WriteLine](writeline-method.md)|Writes a specified text and a new-line character to a TextStream file. |

## Properties

|Property|Description|
|:-------|:----------|
|[AtEndOfLine](atendofline-property.md)|Returns true if the file pointer is positioned immediately before the end-of-line marker in a TextStream file, and false if not. |
|[AtEndOfStream](atendofstream-property.md)|Returns true if the file pointer is at the end of a TextStream file, and false if not. |
|[Column](column-property-visual-basic-for-applications.md)|Returns the column number of the current character position in an input stream. |
|[Line](line-property.md)|Returns the current line number in a TextStream file. |


## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
