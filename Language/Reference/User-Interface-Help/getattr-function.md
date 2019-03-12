---
title: GetAttr function (Visual Basic for Applications)
keywords: vblr6.chm1008929
f1_keywords:
- vblr6.chm1008929
ms.prod: office
ms.assetid: e64ff896-0fae-8a77-7b4c-9d21e83ff919
ms.date: 12/12/2018
localization_priority: Normal
---


# GetAttr function

Returns an **Integer** representing the attributes of a file, directory, or folder.

## Syntax

**GetAttr**(_pathname_)

The required _pathname_ [argument](../../Glossary/vbe-glossary.md#argument) is a [string expression](../../Glossary/vbe-glossary.md#string-expression) that specifies a file name. The _pathname_ may include the directory or folder, and the drive.

## Return values

The value returned by **GetAttr** is the sum of the following attribute values:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbNormal**|0|Normal.|
|**vbReadOnly**|1|Read-only.|
|**vbHidden**|2|Hidden.|
|**vbSystem**|4|System file. Not available on the Macintosh.|
|**vbDirectory**|16|Directory or folder.|
|**vbArchive**|32|File has changed since last backup. Not available on the Macintosh.|
|**vbAlias**|64|Specified file name is an alias. Available only on the Macintosh.|

> [!NOTE] 
> These [constants](../../Glossary/vbe-glossary.md#constant) are specified by Visual Basic for Applications. The names can be used anywhere in your code in place of the actual values.

## Remarks

To determine which attributes are set, use the **[And](and-operator.md)** operator to perform a [bitwise comparison](../../Glossary/vbe-glossary.md#bitwise-comparison) of the value returned by the **GetAttr** function and the value of the individual file attribute you want. If the result is not zero, that attribute is set for the named file. For example, the return value of the following **And** expression is zero if the Archive attribute is not set:

```vb
Result = GetAttr(FName) And vbArchive
```

A nonzero value is returned if the Archive attribute is set.

## Example

This example uses the **GetAttr** function to determine the attributes of a file and directory or folder. On the Macintosh, only the constants **vbNormal**, **vbReadOnly**, **vbHidden**, and **vbAlias** are available.

```vb
Dim MyAttr
' Assume file TESTFILE has hidden attribute set.
MyAttr = GetAttr("TESTFILE")    ' Returns 2.

' Returns nonzero if hidden attribute is set on TESTFILE.
Debug.Print MyAttr And vbHidden    

' Assume file TESTFILE has hidden and read-only attributes set.
MyAttr = GetAttr("TESTFILE")    ' Returns 3.

' Returns nonzero if hidden attribute is set on TESTFILE.
Debug.Print MyAttr And (vbHidden + vbReadOnly)    

' Assume MYDIR is a directory or folder.
MyAttr = GetAttr("MYDIR")    ' Returns 16.

```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
