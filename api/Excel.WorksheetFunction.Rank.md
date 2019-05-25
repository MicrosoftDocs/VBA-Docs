---
title: WorksheetFunction.Rank method (Excel)
keywords: vbaxl10.chm137159
f1_keywords:
- vbaxl10.chm137159
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rank
ms.assetid: e75cabc4-1d97-b8fd-4e7d-3b12ab6a53c5
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Rank method (Excel)

Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list. If you were to sort the list, the rank of the number would be its position.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[Rank_Eq](Excel.WorksheetFunction.Rank_Eq.md)** and **[Rank_Avg](Excel.WorksheetFunction.Rank_Avg.md)** methods.


## Syntax

_expression_.**Rank** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the number whose rank you want to find.|
| _Arg2_|Required| **Range**|Ref - an array of, or a reference to, a list of numbers. Nonnumeric values in ref are ignored.|
| _Arg3_|Optional| **Variant**|Order - a number specifying how to rank number.|

## Return value

**Double**


## Remarks

If order is 0 (zero) or omitted, Microsoft Excel ranks number as if ref were a list sorted in descending order.
    
If order is any nonzero value, Excel ranks number as if ref were a list sorted in ascending order.
    
**Rank** gives duplicate numbers the same rank. However, the presence of duplicate numbers affects the ranks of subsequent numbers. For example, in a list of integers sorted in ascending order, if the number 10 appears twice and has a rank of 5, 11 would have a rank of 7 (no number would have a rank of 6). 

For some purposes you might want to use a definition of rank that takes ties into account. In the previous example, you would want a revised rank of 5.5 for the number 10. To do this, add the following correction factor to the value returned by **Rank**. This correction factor is appropriate both for the case where rank is computed in descending order (order = 0 or omitted) or ascending order (order = nonzero value). 

- Correction factor for tied ranks =[COUNT(ref) + 1 – RANK(number, ref, 0) – RANK(number, ref, 1)]/2. 

- In the following example, RANK(A2,A1:A5,1) equals 3. The correction factor is (5 + 1 – 2 – 3)/2 = 0.5, and the revised rank that takes ties into account is 3 + 0.5 = 3.5. 

- If number occurs only once in ref, the correction factor will be 0 because **Rank** would not have to be adjusted for a tie.


    

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
