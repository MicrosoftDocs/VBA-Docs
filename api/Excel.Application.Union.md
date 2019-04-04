---
title: Application.Union method (Excel)
keywords: vbaxl10.chm132112
f1_keywords:
- vbaxl10.chm132112
ms.prod: excel
api_name:
- Excel.Application.Union
ms.assetid: 7c70a5be-2696-5fc2-bd69-6c6ff4d3291e
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Union method (Excel)

Returns the union of two or more ranges.


## Syntax

_expression_.**Union** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|At least two **[Range](Excel.Range(object).md)** objects must be specified.|
| _Arg2_|Required| **Range**|At least two **Range** objects must be specified.|
| _Arg3_ &ndash; _Arg30_ |Optional| **Variant**|A range.|


## Return value

Range


## Example

This example fills the union of two named ranges, Range1 and Range2, with the formula =RAND().

```vb
Worksheets("Sheet1").Activate 
Set bigRange = Application.Union(Range("Range1"), Range("Range2")) 
bigRange.Formula = "=RAND()"
```

<br/>

This example compares the **[Worksheet.Range](Excel.Worksheet.Range.md)** property, **Application.Union** method, and **[Application.Intersect](Excel.Application.Intersect.md)** method.

 ```vb
Range("A1:A10").Select                            'Selects cells A1 to A10.
Range(Range("A1"), Range("A10")).Select           'Selects cells A1 to A10.
 Range("A1, A10").Select                          'Selects cells A1 and A10.
Union(Range("A1"), Range("A10")).Select           'Selects cells A1 and A10.
 Range("A1:A5 A5:A10").Select                     'Selects cell A5.
Intersect(Range("A1:A5"), Range("A5:A10")).Select 'Selects cell A5.
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
