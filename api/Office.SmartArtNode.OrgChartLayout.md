---
title: SmartArtNode.OrgChartLayout property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.OrgChartLayout
ms.assetid: 183879a1-94fe-e102-51ec-66146d002f75
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartArtNode.OrgChartLayout property (Office)

Retrieves or sets the **[MsoOrgChartLayoutType](office.msoorgchartlayouttype.md)** associated with this node if there is one. Read/write.


## Syntax

_expression_.**OrgChartLayout**

_expression_ An expression that returns a **[SmartArtNode](Office.SmartArtNode.md)** object.


## Remarks

Possible members are:

- msoOrgChartLayoutBothHanging
    
- msoOrgChartLayoutDefault
    
- msoOrgChartLayoutLeftHanging
    
- msoOrgChartLayoutMixed
    
- msoOrgChartLayoutRightHanging
    
- msoOrgChartLayoutStandard
    

## Example

The following code sets the **OrgChartLayout** property to the default layout.

```vb
Dim saNode As SmartArtNode 
saNode.OrgChartLayout = msoOrgChartLayoutDefault
```


## See also

- [SmartArtNode object members](overview/Library-Reference/smartartnode-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]