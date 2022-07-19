---
title: Split function (Visual Basic for Applications)
keywords: vblr6.chm1008907
f1_keywords:
- vblr6.chm1008907
ms.prod: office
ms.assetid: 7c68f50a-c4c4-ee16-cc04-9d067a0b5819
ms.date: 12/13/2018
ms.localizationpriority: medium
---


# Split function

Returns a zero-based, one-dimensional [array](../../Glossary/vbe-glossary.md#array) containing a specified number of substrings.

## Syntax

**Split**(_expression_, [ _delimiter_, [ _limit_, [ _compare_ ]]])

The **Split** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_expression_|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) containing substrings and delimiters. If _expression_ is a zero-length string(""), **Split** returns an empty array, that is, an array with no elements and no data.|
|_delimiter_|Optional. String character used to identify substring limits. If omitted, the space character (" ") is assumed to be the delimiter. If _delimiter_ is a zero-length string, a single-element array containing the entire _expression_ string is returned.|
|_limit_|Optional. Number of substrings to be returned; -1 indicates that all substrings are returned.|
|_compare_|Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.|

## Settings

The _compare_ argument can have the following values:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison by using the setting of the **[Option Compare](option-compare-statement.md)** statement.|
|**vbBinaryCompare**|0|Performs a binary comparison.|
|**vbTextCompare**|1|Performs a textual comparison.|
|**vbDatabaseCompare**|2|Microsoft Access only. Performs a comparison based on information in your database.|

## Example

This example shows how to use the **Split** function. 

```vb
Dim strFull As String
Dim arrSplitStrings1() As String
Dim arrSplitStrings2() As String
Dim strSingleString1 As String
Dim strSingleString2 As String
Dim strSingleString3 As String
Dim i As Long

strFull = "Dow - Fonseca - Graham - Kopke - Noval - Offley - Sandeman - Taylor - Warre"    ' String that will be used. 

arrSplitStrings1 = Split(strFull, "-")      ' arrSplitStrings1 will be an array from 0 To 8. 
                                            ' arrSplitStrings1(0) = "Dow " and arrSplitStrings1(1) = " Fonesca ". 
                                            ' The delimiter did not include spaces, so the spaces in strFull will be included in the returned array values. 

arrSplitStrings2 = Split(strFull, " - ")    ' arrSplitStrings2 will be an array from 0 To 8. 
                                            ' arrSplitStrings2(0) = "Dow" and arrSplitStrings2(1) = "Fonesca". 
                                            ' The delimiter includes the spaces, so the spaces will not be included in the returned array values. 

'Multiple examples of how to return the value "Kopke" (array position 3). 

strSingleString1 = arrSplitStrings2(3)      ' strSingleString1 = "Kopke". 

strSingleString2 = Split(strFull, " - ")(3) ' strSingleString2 = "Kopke".
                                            ' This syntax can be used if the entire array is not needed, and the position in the returned array for the desired value is known. 

For i = LBound(arrSplitStrings2, 1) To UBound(arrSplitStrings2, 1)
    If InStr(1, arrSplitStrings2(i), "Kopke", vbTextCompare) > 0 Then
        strSingleString3 = arrSplitStrings2(i)
        Exit For
    End If 
Next i

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
