---
title: Print statement (VBA)
keywords: vblr6.chm1008995
f1_keywords:
- vblr6.chm1008995
ms.prod: office
ms.assetid: 47c69cf9-2476-b9c2-782c-1c0fc2747936
ms.date: 12/03/2018
localization_priority: Normal
---


# Print # statement

Writes display-formatted data to a sequential file.

## Syntax

**Print** **#**_filenumber_, [ _outputlist_ ]

<br/>

The **Print #** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](../../Glossary/vbe-glossary.md#file-number).|
| _outputlist_|Optional. [Expression](../../Glossary/vbe-glossary.md#expression) or list of expressions to print.|

## Settings

The _outputlist_ [argument](../../Glossary/vbe-glossary.md#argument) settings are:

[{ **Spc**(_n_) | **Tab** [ (_n_) ]}] [ _expression_ ] [ _charpos_ ]

<br/>

|Setting|Description|
|:-----|:-----|
|**Spc**(_n_) |Used to insert space characters in the output, where _n_ is the number of space characters to insert.|
|**Tab**(_n_) |Used to position the insertion point to an absolute column number, where _n_ is the column number. Use **Tab** with no argument to position the insertion point at the beginning of the next [print zone](../../Glossary/vbe-glossary.md#print-zone).|
| _expression_|[Numeric expressions](../../Glossary/vbe-glossary.md#numeric-expression) or [string expressions](../../Glossary/vbe-glossary.md#string-expression) to print.|
| _charpos_|Specifies the insertion point for the next character. Use a semicolon to position the insertion point immediately after the last character displayed. Use **Tab**(_n_) to position the insertion point to an absolute column number. Use **Tab** with no argument to position the insertion point at the beginning of the next print zone. If _charpos_ is omitted, the next character is printed on the next line.|

## Remarks

Data written with **Print #** is usually read from a file with **[Line Input #](line-inputstatement.md)** or **[Input #](inputstatement.md)**.

If you omit _outputlist_ and include only a list separator after _filenumber_, a blank line is printed to the file.

Multiple expressions can be separated with either a space or a semicolon. A space has the same effect as a semicolon.

For [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) data, either `True` or `False` is printed. The **True** and **False** keywords are not translated, regardless of the [locale](../../Glossary/vbe-glossary.md#locale).

[Date](../../Glossary/vbe-glossary.md#date-data-type) data is written to the file by using the standard short date format recognized by your system. When either the date or the time component is missing or zero, only the part provided gets written to the file.

Nothing is written to the file if _outputlist_ data is [Empty](../../Glossary/vbe-glossary.md#empty). However, if _outputlist_ data is [Null](../../Glossary/vbe-glossary.md#null), **Null** is written to the file.

For **Error** data, the output appears as `Error` _errorcode_. The **Error** keyword is not translated regardless of the locale.

All data written to the file by using **Print #** is internationally-aware; that is, the data is properly formatted by using the appropriate decimal separator.

Because **Print #** writes an image of the data to the file, you must delimit the data so that it prints correctly. If you use **Tab** with no arguments to move the print position to the next print zone, **Print #** also writes the spaces between print fields to the file.

> [!NOTE] 
> If, at some future time, you want to read the data from a file by using the **Input #** statement, use the **[Write #](writestatement.md)** statement instead of the **Print #** statement to write the data to the file. Using **Write #** ensures the integrity of each separate data field by properly delimiting it, so that it can be read back in by using **Input #**. Using **Write #** also ensures that it can be correctly read in any locale.


## Example

This example uses the **Print #** statement to write data to a file.


```vb
Open "TESTFILE" For Output As #1 ' Open file for output. 
Print #1, "This is a test" ' Print text to file. 
Print #1, ' Print blank line to file. 
Print #1, "Zone 1"; Tab ; "Zone 2" ' Print in two print zones. 
Print #1, "Hello" ; " " ; "World" ' Separate strings with space. 
Print #1, Spc(5) ; "5 leading spaces " ' Print five leading spaces. 
Print #1, Tab(10) ; "Hello" ' Print word at column 10. 
 
' Assign Boolean, Date, Null and Error values. 
Dim MyBool, MyDate, MyNull, MyError 
MyBool = False : MyDate = #February 12, 1969# : MyNull = Null 
MyError = CVErr(32767) 
' True, False, Null, and Error are translated using locale settings of 
' your system. Date literals are written using standard short date 
' format. 
Print #1, MyBool ; " is a Boolean value" 
Print #1, MyDate ; " is a date" 
Print #1, MyNull ; " is a null value" 
Print #1, MyError ; " is an error value" 
Close #1 ' Close file. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
