---
title: FormatCondition object (Excel)
keywords: vbaxl10.chm511072
f1_keywords:
- vbaxl10.chm511072
ms.prod: excel
api_name:
- Excel.FormatCondition
ms.assetid: 38a2bca9-9b28-3ef2-8c7a-4d35a27229ec
ms.date: 07/18/2019
localization_priority: Normal
---


# FormatCondition object (Excel)

Represents a conditional format.


## Remarks

The **FormatCondition** object is a member of the **[FormatConditions](Excel.FormatConditions.md)** collection. The **FormatConditions** collection can now contain more than three conditional formats for a given range.

Use the **[Add](Excel.FormatConditions.Add.md)** method of the **FormatConditions** object to create a new conditional format. If a range has multiple formats, you can use the **Modify** method to change one of the formats, or you can use the **Delete** method to delete a format, and then use the **Add** method to create a new format.

Use the **Font**, **Borders**, and **Interior** properties of the **FormatCondition** object to control the appearance of formatted cells. Some properties of these objects aren't supported by the conditional format object model. Some of the properties that can be used with conditional formatting are listed in the following table.

|Object|Properties you can use with conditional formatting|
|:-----|:-------------------------------------------------|
|**[Font](Excel.Font(object).md)**|**Bold**, **Color**, **ColorIndex**, **FontStyle**, **Italic**, **Strikethrough**, **ThemeColor**, **ThemeFont**, **TintAndShade**, and **Underline**<br/><br/>The accounting underline styles cannot be used.|
|**[Borders](Excel.Borders.md)**|The following borders can be used (all others aren't supported): **xlBottom**, **xlLeft**, **xlRight**, **xlTop**.<br/><br/>The following border styles can be used (all others aren't supported): **xlLineStyleNone**, **xlContinuous**, **xlDash**, **xlDot**, **xlDashDot**, **xlDashDotDot**, **xlGray50**, **xlGray75**, and **xlGray25**.<br/><br/>The following border weights can be used (all others aren't supported): **xlHairline** and **xlThin**.|
|**[Interior](Excel.Interior(object).md)**|**Color**, **ColorIndex**, **Gradient**, **Pattern**, **PatternColor**, **PatternColorIndex**, **PatternThemeColor**, **PatternTintAndShade**, **ThemeColor**, and **TintAndShade**.|

## Example

Use **[FormatConditions](Excel.Range.FormatConditions.md)** (_index_), where _index_ is the index number of the conditional format, to return a **FormatCondition** object. The following example sets format properties for an existing conditional format for cells E1:E10.

```vb
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
 With .Font 
 .Bold = True 
 .ColorIndex = 3 
 End With 
End With
```


## Methods

- [Delete](Excel.FormatCondition.Delete.md)
- [Modify](Excel.FormatCondition.Modify.md)
- [ModifyAppliesToRange](Excel.FormatCondition.ModifyAppliesToRange.md)
- [SetFirstPriority](Excel.FormatCondition.SetFirstPriority.md)
- [SetLastPriority](Excel.FormatCondition.SetLastPriority.md)

## Properties

- [Application](Excel.FormatCondition.Application.md)
- [AppliesTo](Excel.FormatCondition.AppliesTo.md)
- [Borders](Excel.FormatCondition.Borders.md)
- [Creator](Excel.FormatCondition.Creator.md)
- [DateOperator](Excel.FormatCondition.DateOperator.md)
- [Font](Excel.FormatCondition.Font.md)
- [Formula1](Excel.FormatCondition.Formula1.md)
- [Formula2](Excel.FormatCondition.Formula2.md)
- [Interior](Excel.FormatCondition.Interior.md)
- [NumberFormat](Excel.FormatCondition.NumberFormat.md)
- [Operator](Excel.FormatCondition.Operator.md)
- [Parent](Excel.FormatCondition.Parent.md)
- [Priority](Excel.FormatCondition.Priority.md)
- [PTCondition](Excel.FormatCondition.PTCondition.md)
- [ScopeType](Excel.FormatCondition.ScopeType.md)
- [StopIfTrue](Excel.FormatCondition.StopIfTrue.md)
- [Text](Excel.FormatCondition.Text.md)
- [TextOperator](Excel.FormatCondition.TextOperator.md)
- [Type](Excel.FormatCondition.Type.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
