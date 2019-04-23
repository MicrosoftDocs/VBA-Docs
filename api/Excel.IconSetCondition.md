---
title: IconSetCondition object (Excel)
keywords: vbaxl10.chm811072
f1_keywords:
- vbaxl10.chm811072
ms.prod: excel
api_name:
- Excel.IconSetCondition
ms.assetid: e3c4ef69-4d95-87c9-5059-805775288e24
ms.date: 03/30/2019
localization_priority: Normal
---


# IconSetCondition object (Excel)

Represents an icon set conditional formatting rule.


## Remarks

All conditional formatting objects are contained within a **[FormatConditions](Excel.FormatConditions.md)** collection object, which is a child of a **[Range](Excel.Range(object).md)** collection. 

You can create an icon set formatting rule by using either the **[Add](Excel.FormatConditions.Add.md)** method or **[AddIconSetCondition](Excel.FormatConditions.AddIconSetCondition.md)** method of the **FormatConditions** collection.

Each icon set contains three, four, or five icons. You use the **[IconSets](Excel.Workbook.IconSets.md)** property of the **Workbook** object to return an **[IconSets](Excel.IconSets.md)** object to specify one of the built-in icon sets. Each individual icon in the icon set is then assigned to a subset of the values of the range by the members of the **[IconCriteria](Excel.IconCriteria.md)** object. The type of threshold is also specified by this object.


## Example

The following code example creates a range of numbers representing test scores, and then applies an icon set conditional formatting rule to that range. The type of icon set is then changed from the default icons to a five-arrow icon set. Finally, the threshold type is modified from percentile to a hard-coded number.

```vb
Sub CreateIconSetCF() 
 
    Dim cfIconSet As IconSetCondition 
     
    'Fill cells with sample data from 1 to 10 
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

## Methods

- [Delete](Excel.IconSetCondition.Delete.md)
- [ModifyAppliesToRange](Excel.IconSetCondition.ModifyAppliesToRange.md)
- [SetFirstPriority](Excel.IconSetCondition.SetFirstPriority.md)
- [SetLastPriority](Excel.IconSetCondition.SetLastPriority.md)

## Properties

- [Application](Excel.IconSetCondition.Application.md)
- [AppliesTo](Excel.IconSetCondition.AppliesTo.md)
- [Creator](Excel.IconSetCondition.Creator.md)
- [Formula](Excel.IconSetCondition.Formula.md)
- [IconCriteria](Excel.IconSetCondition.IconCriteria.md)
- [IconSet](Excel.IconSetCondition.IconSet.md)
- [Parent](Excel.IconSetCondition.Parent.md)
- [PercentileValues](Excel.IconSetCondition.PercentileValues.md)
- [Priority](Excel.IconSetCondition.Priority.md)
- [PTCondition](Excel.IconSetCondition.PTCondition.md)
- [ReverseOrder](Excel.IconSetCondition.ReverseOrder.md)
- [ScopeType](Excel.IconSetCondition.ScopeType.md)
- [ShowIconOnly](Excel.IconSetCondition.ShowIconOnly.md)
- [StopIfTrue](Excel.IconSetCondition.StopIfTrue.md)
- [Type](Excel.IconSetCondition.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]