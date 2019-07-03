---
title: Cell Formulas 
keywords: ???
f1_keywords: ???
ms.prod: excel
ms.assetid: ??? 
ms.date: ???
localization_priority: Normal
---

# Formula Variations

Range.Formula and Range.Formula2 are two different ways of representing the logic in the formula. They can be throught of a 2 dialects of Excel's formula language.

Excel has always supported 2 types of formula evaluation: Implicitly Intersection Evaluation (IIE) and Array Evaluation (AE). IIE was the default for cell formulas, while AE was used everywhere else (Conditional Formatting, Data Validation, CSE Array formulas, etc).  

The primary difference between the two forms of Evaluation was how they behaved when an array (e.g. {1,2,3})or multi celled range (e.g. A1:A10) was returned as the result of the formula or a range was passed to a function that expected a single value:

* IIE would choose one value. For an array it would return the top left value. For a multi cell range it would select a cell on the same row or column as the formula. This operation is referred to as "implicit intersection".
* AE would use all the values, calling the function multiple times and return an array results. This operation is referred to as "lifting".

IIE was the default for cell formulas as it guarenteed that only one value would be returned. Cell formulas that used IIE are set using Range.Formula.

With the introduction of Dyanamic Arrays (DA), Excel now supports returning multiple values to the grid and AE is now the default. AE formula's can be set/read using Range.Formula2 which supersedes Range.Formula. However, to facilitate backcompatiblity, Range.Formula is still supported and will continue to set/return IIE formulas. Formula's set using Range.Formula will trigger implicit intersection and can never spill. Formula read using Range.Formula will be silent on where Implicit Intersection occurs.

Range.Formula can be thought of as the formula would be presented in the formula bar in Pre-DA Excel, while Range.Formula2 is how the formula will be presented in DA Excel.

Excel automatically translates between these two formula variations, so either can be read and set. To facilitate the translation from Range.Formula (using IIE) to Range.Formula2 (AE), Excel will indicate where previously implicit intersection could occur using the new implicit intersection operator @. Likewise, to facilitate the translation from Range.Formula2 (using AE) to Range.Formula (using IIE) Excel will remove @ operators that would be performed silently. Often there is no difference between the two because many common formulas do provide arrays or multicelled ranges to functions that expect single values or return multicelled ranges as their result which is the only time IIE and AE behave differently.

# Translating from Range.Formula to Range.Formula2

This example shows the outcome of setting Range.Formula and then retrieving Range.Formula2

```vb
Dim cell As Range
Dim str As String

Set cell = Worksheets("Sheet1").Cells(2, 1)
ArrayOfFormulas = Array("=SQRT(A1)", "=SQRT(A1:A4)")

For i = LBound(ArrayOfFormulas) To UBound(ArrayOfFormulas)
 cell.Formula = ArrayOfFormulas(i)
 str = "Wrote Range.Formula:" & vbCr & cell.Formula & vbCr & vbCr & "Read Range.Formula2:" & vbCr & cell.Formula2
 MsgBox (str)
Next i
```

|Write Range.Formula|Read Range.Formula2|Notes|
|---|---|---|
|=SQRT(A1)|=SQRT(A1)|Identical because no implicit intersection could occur|
|=SQRT(A1:A4)|=SQRT(@A1:A4)|SQRT expects a single value but is given an multi celled range. This will trigger implicit intersection in IIE, therefor the translation to AE calls out where implicit intersection could occur using the @ operator|


#Translating from Range.Formula2 to Range.Formula

Formula set using Range.Formula2 Excel will ensure that is calculated using AE. This is done by evaluating all newly authored formulas as array formulas. However, to minimize the number of CSE array formulas shown in old Excel, DA Excel analyses the formula to determine if an array or mulitcelled range could be provided to a function that expected a single value - the only time IIE and AE differ. If the formula will calcs the same under IIE and AE, DA Excel saves the formula as an IIE. If there is any potential that they differ, Excel will it to file in save such that pre-DA excel sees it as an array formula. You can test whether the formula will appear as an array formula for pre-DA Excel using Range.IsSavedAsArray()


```vb
Dim cell As Range
Dim str As String

Set cell = Worksheets("Sheet1").Cells(2, 1)
ArrayOfFormulas = Array("=SQRT(A1)", "=SQRT(@A1:A4)", "=SQRT(A1:A4)", "=SQRT(A1:A4)+SQRT(@A1:A4)")

For i = LBound(ArrayOfFormulas) To UBound(ArrayOfFormulas)
 cell.Formula2 = ArrayOfFormulas(i)
 str = "Wrote Range.Formula2:" & vbCr & cell.Formula2 & vbCr & "Read Range.Formula:" & vbCr & cell.Formula & vbCr & "Read Range.IsSavedAsArray:" & vbCr & cell.SavedAsArray
 MsgBox (str)
Next i
```

CHANGE EXAMPLES TO USE FUNCTIONS. DISCUSS HOW FUNCTION EXPECTS single value

|**Write Range.Formula2**|**Read Range.Formula**|**Read Range.isSavedAsArray**|**Notes**|
|---|---|---|---|
|=SQRT(A1)|=SQRT(A1)|FALSE|SQRT expects a single value, A1 is a single value. Therefor no variance between IIE and AE. Save as IIE and remove any @'s|
|=SQRT(@A1:A4)|=SQRT(A1:A4)|FALSE|SQRT expects a single value, @A1:A4 is a single value. Therefor no variance between IIE and AE. Save as IIE and remove any @'s|
|=SQRT(A1:A4)|=SQRT(A1:A4)|TRUE|SQRT expects a single value, A1:A4 is a multicell range. IIE and AE could vary therefor save as array|
|=SQRT(A1:A4)+SQRT(@A1:A4)|=SQRT(A1:A4)+SQRT(@A1:A4)|TRUE|The first SQRT expects a single value, A1:A4 is a multicell range. IIE and AE could vary therefor save as array|

# Best Practice

If targeting DA version of Excel, you should use Range.Formula2 in preference to Range.Formula.

If targeting Pre and Post DA version of Excel, you should continue to use Range.Formula. If however you want tight control over the appearance of the formula the users formula bar, you should detect whether .Formula2 is supported and, if so, use .Formula2 otherwise use .Formula 

# Notes

OfficeJS does not include Range.Formula2. Instead Range.Formula always reports what is present in the formula bar. As a newer language with the ability for addins to quickly deploy updates, developers are encouraged to update their addins if they encounter any compatibility issues between AE to IIE.    

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
