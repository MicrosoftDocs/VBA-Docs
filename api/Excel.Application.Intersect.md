---
title: Application.Intersect method (Excel)
keywords: vbaxl10.chm183099
f1_keywords:
- vbaxl10.chm183099
ms.prod: excel
api_name:
- Excel.Application.Intersect
ms.assetid: 856d052a-3207-ced2-941c-b466cb880a93
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Intersect method (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the rectangular intersection of two or more ranges. If one or more ranges from a different worksheet are specified, an error is returned.


## Syntax

_expression_.**Intersect** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters


|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|The intersecting ranges. At least two **Range** objects must be specified.|
| _Arg2_|Required| **Range**|The intersecting ranges. At least two **Range** objects must be specified.|
| _Arg3_&ndash;_Arg30_|Optional| **Variant**|An intersecting range.|



## Return value

Range


## Example

The following example selects the intersection of two named ranges, rg1 and rg2, on Sheet1. If the ranges don't intersect, the example displays a message.

```vb
Worksheets("Sheet1").Activate 
Set isect = Application.Intersect(Range("rg1"), Range("rg2")) 
If isect Is Nothing Then 
 MsgBox "Ranges do not intersect" 
Else 
 isect.Select 
End If
```

<br/>

The following example compares the **[Worksheet.Range](Excel.Worksheet.Range.md)** property, the **[Application.Union](Excel.Application.Union.md)** method, and the **Intersect** method.

 ```vb
Range("A1:A10").Select                            'Selects cells A1 to A10.
Range(Range("A1"), Range("A10")).Select           'Selects cells A1 to A10.
 Range("A1, A10").Select                           'Selects cells A1 and A10.
Union(Range("A1"), Range("A10")).Select           'Selects cells A1 and A10.
 Range("A1:A5 A5:A10").Select                      'Selects cell A5.
Intersect(Range("A1:A5"), Range("A5:A10")).Select 'Selects cell A5.
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
