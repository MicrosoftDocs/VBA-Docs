---
title: FormatCondition Object (Excel)
keywords: vbaxl10.chm511072
f1_keywords:
- vbaxl10.chm511072
ms.prod: excel
api_name:
- Excel.FormatCondition
ms.assetid: 38a2bca9-9b28-3ef2-8c7a-4d35a27229ec
ms.date: 06/08/2017
---


# FormatCondition Object (Excel)

Represents a conditional format.


## Remarks

 The **FormatCondition** object is a member of the **[FormatConditions](Excel.FormatConditions.md)** collection. The **FormatConditions** collection can now contain more than three conditional formats for a given range.

Use the  **[Add](Excel.FormatConditions.Add.md)** method to create a new conditional format. If a range has mulitple formats, you can use the **[Modify](Excel.FormatCondition.Modify.md)** method to change one of the formats, or you can use the **[Delete](Excel.FormatCondition.Delete.md)** method to delete a format and then use the **Add** method to create a new format.

Use the  **[Font](Excel.FormatCondition.Font.md)**, **[Borders](Excel.FormatCondition.Borders.md)**, and **[Interior](Excel.FormatCondition.Interior.md)** properties of the **FormatCondition** object to control the appearance of formatted cells. Some properties of these objects aren?t supported by the conditional format object model. Some of the properties that can be used with conditional formatting are listed in the following table.



|**Object**|**Properties**|
|:-----|:-----|
|**[Font](Excel.Font(objec).md)**|**Bold** **Color** **ColorIndex** **FontStyle** **Italic** **Strikethrough** **Underline** The accounting underline styles cannot be used.|
|**[Border](Excel.Border(objec).md)**|**Bottom** **Color** **Left** **Right** **Style** The following border styles can be used (all others aren?t supported): **xlNone**, **xlSolid**, **xlDash**, **xlDot**, **xlDashDot**, **xlDashDotDot**, **xlGray50**, **xlGray75**, and **xlGray25**. **Top** **Weight** The following border weights can be used (all others aren?t supported): **xlWeightHairline** and **xlWeightThin**.|
|**[Interior](Excel.Interior(objec).md)**|**Color** **ColorIndex** **Pattern** **PatternColorIndex**|

## Example

Use  **[FormatConditions](Excel.Range.FormatConditions.md)** ( _index_ ), where _index_ is the index number of the conditional format, to return a **FormatCondition** object. The following example sets format properties for an existing conditional format for cells E1:E10.


```
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



|**Name**|
|:-----|
|[Delete](Excel.FormatCondition.Delete.md)|
|[Modify](Excel.FormatCondition.Modify.md)|
|[ModifyAppliesToRange](Excel.FormatCondition.ModifyAppliesToRange.md)|
|[SetFirstPriority](Excel.FormatCondition.SetFirstPriority.md)|
|[SetLastPriority](Excel.FormatCondition.SetLastPriority.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.FormatCondition.Application.md)|
|[AppliesTo](Excel.FormatCondition.AppliesTo.md)|
|[Borders](Excel.FormatCondition.Borders.md)|
|[Creator](Excel.FormatCondition.Creator.md)|
|[DateOperator](Excel.FormatCondition.DateOperator.md)|
|[Font](Excel.FormatCondition.Font.md)|
|[Formula1](Excel.FormatCondition.Formula1.md)|
|[Formula2](Excel.FormatCondition.Formula2.md)|
|[Interior](Excel.FormatCondition.Interior.md)|
|[NumberFormat](Excel.FormatCondition.NumberFormat.md)|
|[Operator](Excel.FormatCondition.Operator.md)|
|[Parent](Excel.FormatCondition.Parent.md)|
|[Priority](Excel.FormatCondition.Priority.md)|
|[PTCondition](Excel.FormatCondition.PTCondition.md)|
|[ScopeType](formatcondition-scopetype-property-excel.md)|
|[StopIfTrue](formatcondition-stopiftrue-property-excel.md)|
|[Text](formatcondition-text-property-excel.md)|
|[TextOperator](formatcondition-textoperator-property-excel.md)|
|[Type](formatcondition-type-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
