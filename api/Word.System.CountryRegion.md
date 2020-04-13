---
title: System.CountryRegion property (Word)
keywords: vbawd10.chm154468455
f1_keywords:
- vbawd10.chm154468455
ms.prod: word
api_name:
- Word.System.CountryRegion
ms.assetid: 51db26e6-9f24-5934-24a4-0ed87bb51f69
ms.date: 06/08/2017
localization_priority: Normal
---


# System.CountryRegion property (Word)

Returns the country/region designation of the system. Read-only  **WdCountry**.


## Syntax

_expression_. `CountryRegion`

_expression_ Required. A variable that represents a '[System](Word.System.md)' object.


## Example

If the **CountryRegion** property returns **wdUS**, this example converts the top margin value from points to inches.


```vb
Dim sngMargin As Single 
 
If System.CountryRegion = wdUS Then 
 sngMargin = ActiveDocument.PageSetup.TopMargin 
 MsgBox "Top margin is " & PointsToInches(sngMargin) 
End If
```


## See also


[System Object](Word.System.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]