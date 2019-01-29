---
title: TextRange2.MathZones property (Office)
ms.prod: office
api_name:
- Office.TextRange2.MathZones
ms.assetid: 277aa819-d717-e2f5-5bc7-607abfce20a4
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2.MathZones property (Office)

Sets the starting point and length of a math zone within a text range. Read-only.


## Syntax

_expression_.**MathZones** (_Start_, _Length_)

_expression_ An expression that returns a **[TextRange2](Office.TextRange2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Integer**|The starting point for the math zone.|
| _Length_|Optional|**Integer**|The length of the math zone.|

## Remarks

A math zone is a text range within which math typography rules apply and outside of which math typography rules do not apply. In addition to containing special mathematical symbols, math zones can also contain text such as in the equation `rate = distance/time` where text appears with math symbols.


## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]