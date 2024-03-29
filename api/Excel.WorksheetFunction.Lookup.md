---
title: WorksheetFunction.Lookup method (Excel)
keywords: vbaxl10.chm137089
f1_keywords:
- vbaxl10.chm137089
api_name:
- Excel.WorksheetFunction.Lookup
ms.assetid: 0088c289-2ef5-78ea-68e2-1b10d077e775
ms.date: 05/24/2019
ms.localizationpriority: medium
---


# WorksheetFunction.Lookup method (Excel)

Returns a value either from a one-row or one-column range or from an array. The **Lookup** function has two syntax forms: the vector form and the array form.


## Syntax

_expression_.**Lookup** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Lookup_value - A value that **Lookup** searches for in the first vector. Lookup_value can be a number, text, a logical value, or a name or reference that refers to a value.|
| _Arg2_|Required| **Variant**|Lookup_vector or Array - In vector form, a range that contains only one row or one column. The values in lookup_vector can be text, numbers, or logical values. In array form, a range of cells that contains text, numbers, or logical values that you want to compare with lookup_value.|
| _Arg3_|Optional| **Variant**|Result_vector - Only used with the vector form. A range that contains only one row or column. It must be the same size as lookup_vector.|

## Return value

**Variant**


## Remarks

|If you want to|See|Usage|
|:-----|:-----|:-----|
|Look in a one-row or one-column range (known as a vector) for a value and return a value from the same position in a second one-row or one-column range|Vector form|Use the vector form when you have a large list of values to look up or when the values may change over time.|
|Look in the first row or column of an array for the specified value and return a value from the same position in the last row or column of the array|Array form|Use the array form when you have a small list of values and the values remain constant over time.|

### Vector form

A vector is a range of only one row or one column. The vector form of **Lookup** looks in a one-row or one-column range (known as a vector) for a value and returns a value from the same position in a second one-row or one-column range. Use this form of the **Lookup** function when you want to specify the range that contains the values that you want to match. The other form of **Lookup** automatically looks in the first column or row. 

> [!IMPORTANT] 
> The values in lookup_vector must be placed in ascending order: ...,-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE; otherwise, **Lookup** may not give the correct value. Uppercase and lowercase text are equivalent.

If **Lookup** can't find the lookup_value, it matches the largest value in lookup_vector that is less than or equal to lookup_value.
    
If lookup_value is smaller than the smallest value in lookup_vector, **Lookup** gives the #N/A error value.
    

### Array form

The array form of **Lookup** looks in the first row or column of an array for the specified value and returns a value from the same position in the last row or column of the array. Use this form of **Lookup** when the values that you want to match are in the first row or column of the array. Use the other form of **Lookup** when you want to specify the location of the column or row. 

> [!TIP] 
> In general, it's best to use the **[HLookup](excel.worksheetfunction.hlookup.md)** or **[VLookup](excel.worksheetfunction.vlookup.md)** function instead of the array form of **Lookup**. This form of **Lookup** is provided for compatibility with other spreadsheet programs.

- If **Lookup** can't find the lookup_value, it uses the largest value in the array that is less than or equal to lookup_value. 
    
- If lookup_value is smaller than the smallest value in the first row or column (depending on the array dimensions), **Lookup** returns the #N/A error value. 
    
The array form of **Lookup** is very similar to the **HLookup** and **VLookup** functions. The difference is that **HLookup** searches for lookup_value in the first row, **VLookup** searches in the first column, and **Lookup** searches according to the dimensions of array. 

- If array covers an area that is wider than it is tall (more columns than rows), **Lookup** searches for lookup_value in the first row. 
    
- If array is square or is taller than it is wide (more rows than columns), **Lookup** searches in the first column. 
    
- With **HLookup** and **VLookup**, you can index down or across, but **Lookup** always selects the last value in the row or column. 
    
> [!IMPORTANT] 
> The values in array must be placed in ascending order: ...,-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE; otherwise, **Lookup** may not give the correct value. Uppercase and lowercase text are equivalent.

> [!NOTE] 
> You can also use the **Lookup** function as an alternative to the IF function for elaborate tests or tests for more than seven conditions. See the examples in the array form.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
