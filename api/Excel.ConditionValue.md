---
title: ConditionValue object (Excel)
keywords: vbaxl10.chm803072
f1_keywords:
- vbaxl10.chm803072
ms.prod: excel
api_name:
- Excel.ConditionValue
ms.assetid: a39335db-4e0a-66aa-393b-3aa7e5268c00
ms.date: 03/29/2019
localization_priority: Normal
---


# ConditionValue object (Excel)

Represents how the shortest bar or longest bar is evaluated for a data bar conditional formatting rule.


## Remarks

The **ConditionValue** object is returned by using either the **[MaxPoint](Excel.DataBar.MaxPoint.md)** or **[MinPoint](Excel.DataBar.MinPoint.md)** property of the **DataBar** object.

You can change the type of evaluation from the default setting (lowest value for the shortest bar and highest value for the longest bar) by using the **Modify** method.


## Example

The following example creates a range of data and then applies a data bar to the range. You will notice that because there is an extremely low and high value in the range, the middle values have data bars that are of similar length. To disambiguate the middle values, the sample code uses the **ConditionValue** object to change how the thresholds are evaluated to percentiles.

```vb
Sub CreateDataBarCF() 
 
 Dim cfDataBar As DataBar 
 
 'Create a range of data with a couple of extreme values 
 With ActiveSheet 
 .Range("D1") = 1 
 .Range("D2") = 45 
 .Range("D3") = 50 
 .Range("D2:D3").AutoFill Destination:=Range("D2:D8") 
 .Range("D9") = 500 
 End With 
 
 Range("D1:D9").Select 
 
 'Create a data bar with default behavior 
 Set cfDataBar = Selection.FormatConditions.AddDatabar 
 MsgBox "Because of the extreme values, middle data bars are very similar" 
 
 'The MinPoint and MaxPoint properties return a ConditionValue object 
 'which you can use to change threshold parameters 
 cfDataBar.MinPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=5 
 cfDataBar.MaxPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=75 
 
End Sub
```


## Methods

- [Modify](Excel.ConditionValue.Modify.md)

## Properties

- [Application](Excel.ConditionValue.Application.md)
- [Creator](Excel.ConditionValue.Creator.md)
- [Parent](Excel.ConditionValue.Parent.md)
- [Type](Excel.ConditionValue.Type.md)
- [Value](Excel.ConditionValue.Value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]