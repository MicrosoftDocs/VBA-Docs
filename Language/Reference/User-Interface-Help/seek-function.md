---
title: Seek function (Visual Basic for Applications)
ms.prod: office
ms.assetid: 870aba03-b7ad-c931-928d-33aaf9cf5ab6
ms.date: 12/13/2018
localization_priority: Normal
---


# Seek function

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the current read/write position within a file opened by using the **[Open](open-statement.md)** statement.

## Syntax

**Seek**(_filenumber_)

The required _filenumber_ [argument](../../Glossary/vbe-glossary.md#argument) is an [Integer](../../Glossary/vbe-glossary.md#integer-data-type) containing a valid [file number](../../Glossary/vbe-glossary.md#file-number).

## Remarks

**Seek** returns a value between 1 and 2,147,483,647 (equivalent to 2^31 - 1), inclusive.

The following describes the return values for each file access mode.

|Mode|Return value|
|:-----|:-----|
|**Random**|Number of the next record read or written.|
|**Binary**, **Output**, **Append**, **Input**|Byte position at which the next operation takes place. The first byte in a file is at position 1, the second byte is at position 2, and so on.|

## Example

This example uses the **Seek** function to return the current file position. The example assumes that `TESTFILE` is a file containing records of the user-defined type `Record`.

```vb
Type Record    ' Define user-defined type.
    ID As Integer
    Name As String * 20
End Type
```

<br/>

For files opened in Random mode, **Seek** returns the number of the next record.

```vb
Dim MyRecord As Record    ' Declare variable.
Open "TESTFILE" For Random As #1 Len = Len(MyRecord)
Do While Not EOF(1)    ' Loop until end of file.
    Get #1, , MyRecord    ' Read next record.
    Debug.Print Seek(1)    ' Print record number to the Immediate window.
Loop
Close #1    ' Close file.

```

<br/>

For files opened in modes other than Random mode, **Seek** returns the byte position at which the next operation takes place. Assume that `TESTFILE` is a file containing a few lines of text.

```vb
Dim MyChar
Open "TESTFILE" For Input As #1    ' Open file for reading.
Do While Not EOF(1)    ' Loop until end of file.
    MyChar = Input(1, #1)    ' Read next character of data.
    Debug.Print Seek(1)    ' Print byte position to the Immediate window.
Loop
Close #1    ' Close file.
```


## See also

- [Immediate window](immediate-window.md)
- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
