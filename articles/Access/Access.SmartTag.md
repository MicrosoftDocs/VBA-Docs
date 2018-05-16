---
title: SmartTag Object (Access)
keywords: vbaac10.chm13316
f1_keywords:
- vbaac10.chm13316
ms.prod: access
api_name:
- Access.SmartTag
ms.assetid: ec396ef0-65a4-41bc-ab59-1160e6ef1813
ms.date: 06/08/2017
---


# SmartTag Object (Access)

Represents a smart tag that has been added to a control on a form or report. The  **SmartTag** object is a member of the **[SmartTags](Access.SmartTags.md)** collection.


## Remarks

To return a single  **SmartTag** object, use the **Item** method of the **SmartTags** collection, or use **SmartTags** ( _Index_), where  _Index_ represents the number of the smart tag.


 **Note**  Unlike the  **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0) r`eturns the first smart tag for the specified control.

To return the collection of actions available for the smart tag, use the  **[SmartTagActions](Access.SmartTag.SmartTagActions.md)** property. To perform a smart tag action, use the **[Execute](Access.SmartTagAction.Execute.md)** method of the **[SmartTagAction](Access.SmartTagAction.md)** object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](Access.SmartTag.Delete.md)|Deletes the specified object.|

## Properties



|**Name**|
|:-----|
|[Application](Access.SmartTag.Application.md)|
|[IsMissing](Access.SmartTag.IsMissing.md)|
|[Name](Access.SmartTag.Name.md)|
|[Parent](Access.SmartTag.Parent.md)|
|[Properties](Access.SmartTag.Properties.md)|
|[SmartTagActions](Access.SmartTag.SmartTagActions.md)|
|[XML](smarttag-xml-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
