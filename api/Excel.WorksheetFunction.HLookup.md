---
title: WorksheetFunction.HLookup method (Excel)
keywords: vbaxl10.chm137122
f1_keywords:
- vbaxl10.chm137122
api_name:
- Excel.WorksheetFunction.HLookup
ms.assetid: 6e7b5ad0-3f70-d7a8-b161-ce418107d2a1
ms.date: 05/23/2019
ms.localizationpriority: medium
---


# WorksheetFunction.HLookup method (Excel)

Searches for a value in the top row of a table or an array of values, and then returns a value in the same column from a row that you specify in the table or array. Use **HLookup** when your comparison values are located in a row across the top of a table of data, and you want to look down a specified number of rows. Use **[VLookup](excel.worksheetfunction.vlookup.md)** when your comparison values are located in a column to the left of the data that you want to find.


## Syntax

_expression_.**HLookup** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Lookup_value - the value to be found in the first row of the table. Lookup_value can be a value, a reference, or a text string.|
| _Arg2_|Required| **Variant**|Table_array - a table of information in which data is looked up. Use a reference to a range or a range name.|
| _Arg3_|Required| **Variant**|Row_index_num - the row number in table_array from which the matching value will be returned. A row_index_num of 1 returns the first row value in table_array, a row_index_num of 2 returns the second row value in table_array, and so on. If row_index_num is less than 1, **HLookup** returns the #VALUE! error value; if row_index_num is greater than the number of rows on table_array, **HLookup** returns the #REF! error value.|
| _Arg4_|Optional| **Variant**|Range_lookup - a logical value that specifies whether you want **HLookup** to find an exact match or an approximate match. If **True** or omitted, an approximate match is returned. In other words, if an exact match is not found, the next largest value that is less than lookup_value is returned. If **False**, **HLookup** will find an exact match. If one is not found, the error value #N/A is returned.|

## Return value

**Variant**


## Remarks

If **HLookup** can't find lookup_value, and range_lookup is **True**, it uses the largest value that is less than lookup_value.
    
If lookup_value is smaller than the smallest value in the first row of table_array, **HLookup** returns the #N/A error value.
    
If range_lookup is **False** and lookup_value is text, you can use the wildcard characters, question mark (?) and asterisk (*), in lookup_value. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
