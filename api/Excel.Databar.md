---
title: Databar object (Excel)
keywords: vbaxl10.chm809072
f1_keywords:
- vbaxl10.chm809072
api_name:
- Excel.Databar
ms.assetid: 2684e913-c278-e6be-ba9d-053b6ad58bae
ms.date: 03/29/2019
ms.localizationpriority: medium
---


# Databar object (Excel)

Represents a data bar conditional formating rule. Applying a data bar to a range helps you see the value of a cell relative to other cells.


## Remarks

All conditional formatting objects are contained within a **[FormatConditions](Excel.FormatConditions.md)** collection object, which is a child of a **[Range](Excel.Range(object).md)** collection. You can create a data bar formatting rule by using either the **[Add](Excel.FormatConditions.Add.md)** or **[AddDatabar](Excel.FormatConditions.AddDatabar.md)** methods of the **FormatConditions** collection.

You use the **MinPoint** and **MaxPoint** properties of the **Databar** object to set the values of the shortest bar and longest bar of a range of data. These properties return a **[ConditionValue](Excel.ConditionValue.md)** object, with which you can specify how the thresholds are evaluated.

The **Databar** object also provides properties that enable you to specify an axis line that is displayed when negative values are present, and to specify the color and formatting of data bars.


## Example

The following example creates a range of data, and then applies a data bar to the range. You will notice that because there is an extremely low and high value in the range, the middle values have data bars that are of similar length. To disambiguate the middle values, the sample code uses the **ConditionValue** object to change how the thresholds are evaluated to percentiles.

```vb
Sub CreateDatabarCF() 
 
 Dim cfDatabar As Databar 
 
 ' Create a range of data with a couple of extreme values 
 With ActiveSheet 
 .Range("D1") = 1 
 .Range("D2") = 45 
 .Range("D3") = 50 
 .Range("D2:D3").AutoFill Destination:=Range("D2:D8") 
 .Range("D9") = 500 
 End With 
 
 Range("D1:D9").Select 
 
 ' Create a data bar with default behavior 
 Set cfDatabar = Selection.FormatConditions.AddDatabar 
 MsgBox "Because of the extreme values, middle data bars are very similar" 
 
 ' The MinPoint and MaxPoint properties return a ConditionValue object 
 ' which you can use to change threshold parameters 
 cfDatabar.MinPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=5 
 cfDatabar.MaxPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=75 
 
End Sub
```


## Methods

- [Delete](Excel.Databar.Delete.md)
- [ModifyAppliesToRange](Excel.Databar.ModifyAppliesToRange.md)
- [SetFirstPriority](Excel.Databar.SetFirstPriority.md)
- [SetLastPriority](Excel.Databar.SetLastPriority.md)

## Properties

- [Application](Excel.Databar.Application.md)
- [AppliesTo](Excel.Databar.AppliesTo.md)
- [AxisColor](Excel.Databar.AxisColor.md)
- [AxisPosition](Excel.Databar.AxisPosition.md)
- [BarBorder](Excel.Databar.BarBorder.md)
- [BarColor](Excel.Databar.BarColor.md)
- [BarFillType](Excel.Databar.BarFillType.md)
- [Creator](Excel.Databar.Creator.md)
- [Direction](Excel.Databar.Direction.md)
- [Formula](Excel.Databar.Formula.md)
- [MaxPoint](Excel.Databar.MaxPoint.md)
- [MinPoint](Excel.Databar.MinPoint.md)
- [NegativeBarFormat](Excel.Databar.NegativeBarFormat.md)
- [Parent](Excel.Databar.Parent.md)
- [PercentMax](Excel.Databar.PercentMax.md)
- [PercentMin](Excel.Databar.PercentMin.md)
- [Priority](Excel.Databar.Priority.md)
- [PTCondition](Excel.Databar.PTCondition.md)
- [ScopeType](Excel.Databar.ScopeType.md)
- [ShowValue](Excel.Databar.ShowValue.md)
- [StopIfTrue](Excel.Databar.StopIfTrue.md)
- [Type](Excel.Databar.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]