---
title: Databar object (Excel)
keywords: vbaxl10.chm809072
f1_keywords:
- vbaxl10.chm809072
ms.prod: excel
api_name:
- Excel.Databar
ms.assetid: 2684e913-c278-e6be-ba9d-053b6ad58bae
ms.date: 06/08/2017
localization_priority: Normal
---


# Databar object (Excel)

Represents a data bar conditional formating rule. Applying a data bar to a range helps you see the value of a cell relative to other cells.


## Remarks

All conditional formatting objects are contained within a  **[FormatConditions](Excel.FormatConditions.md)** collection object, which is a child of a **[Range](Excel.Range(object).md)** collection. You can create a data bar formatting rule by using either the **[Add](Excel.FormatConditions.Add.md)** or **[AddDatabar](Excel.FormatConditions.AddDatabar.md)** methods of the **FormatConditions** collection.

You use the  **[MinPoint](Excel.Databar.MinPoint.md)** and **[MaxPoint](Excel.Databar.MaxPoint.md)** properties of the **Databar** object to set the values of the shortest bar and longest bar of a range of data. These properites return a **[ConditionValue](Excel.ConditionValue.md)** object, with which you can specify how the thresholds are evaluated.

The  **Databar** object also provides properties that enable you to specify an axis line that is displayed when negative values are present, and to specify the color and formatting of data bars.


## Example

The following example creates a range of data and then applies a data bar to the range. You will notice that because there is an extremely low and high value in the range, the middle values have data bars that are of similiar length. To disambiguate the middle values, the sample code uses the  **ConditionValue** object to change how the thresholds are evaluated to percentiles.


```vb
Sub CreateDataBarCF() 
 
 Dim cfDataBar As Databar 
 
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
 Set cfDataBar = Selection.FormatConditions.AddDatabar 
 MsgBox "Because of the extreme values, middle data bars are very similar" 
 
 ' The MinPoint and MaxPoint properties return a ConditionValue object 
 ' which you can use to change threshold parameters 
 cfDataBar.MinPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=5 
 cfDataBar.MaxPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=75 
 
End Sub
```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

