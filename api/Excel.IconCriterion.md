---
title: IconCriterion object (Excel)
keywords: vbaxl10.chm814072
f1_keywords:
- vbaxl10.chm814072
ms.prod: excel
api_name:
- Excel.IconCriterion
ms.assetid: 3517d900-4d84-2ded-ccb1-a3d78d3f6c09
ms.date: 03/30/2019
localization_priority: Normal
---


# IconCriterion object (Excel)

Represents the criterion for an individual icon in an icon set. The criterion specifies the range of values and the threshold type associated with the icon in an icon set conditional formatting rule.


## Remarks

All the criteria for an icon set conditional format are contained in an **[IconCriteria](Excel.IconCriteria.md)** collection. You can access each **IconCriterion** object in the collection by passing an index into the collection. See the example for details.


## Example

The following code example creates a range of numbers representing test scores, and then applies an icon set conditional formatting rule to that range. The type of icon set is then changed from the default icons to a five-arrow icon set. Finally, the threshold type is modified from percentile to a hard-coded number.

```vb
Sub CreateIconSetCF() 
 
    Dim cfIconSet As IconSetCondition 
     
    'Fill cells with sample data 
    With ActiveSheet 
        .Range("C1") = 55 
        .Range("C2") = 92 
        .Range("C3") = 88 
        .Range("C4") = 77 
        .Range("C5") = 66 
        .Range("C6") = 93 
        .Range("C7") = 76 
        .Range("C8") = 80 
        .Range("C9") = 79 
        .Range("C10") = 83 
        .Range("C11") = 66 
        .Range("C12") = 74 
    End With 
     
    Range("C1:C12").Select 
       
    'Create an icon set conditional format for the created sample data range 
    Set cfIconSet = Selection.FormatConditions.AddIconSetCondition 
     
    'Change the icon set to a five-arrow icon set 
    cfIconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows) 
     
    'The IconCriterion collection contains all IconCriteria 
    'By indexing into the collection you can modify each criterion 
 
    With cfIconSet.IconCriteria(1) 
        .Type = xlConditionValueNumber 
        .Value = 0 
        .Operator = 7 
    End With 
    With cfIconSet.IconCriteria(2) 
        .Type = xlConditionValueNumber 
        .Value = 60 
        .Operator = 7 
    End With 
    With cfIconSet.IconCriteria(3) 
        .Type = xlConditionValueNumber 
        .Value = 70 
        .Operator = 7 
    End With 
    With cfIconSet.IconCriteria(4) 
        .Type = xlConditionValueNumber 
        .Value = 80 
        .Operator = 7 
    End With 
    With cfIconSet.IconCriteria(5) 
        .Type = xlConditionValueNumber 
        .Value = 90 
        .Operator = 7 
    End With 
         
End Sub
```

## Properties

- [Icon](Excel.IconCriterion.Icon.md)
- [Index](Excel.IconCriterion.Index.md)
- [Operator](Excel.IconCriterion.Operator.md)
- [Type](Excel.IconCriterion.Type.md)
- [Value](Excel.IconCriterion.Value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]