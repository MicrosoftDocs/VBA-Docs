---
title: PropertyPages Object (Outlook)
keywords: vbaol11.chm160
f1_keywords:
- vbaol11.chm160
ms.prod: outlook
api_name:
- Outlook.PropertyPages
ms.assetid: 9850ae7b-f167-d3b2-2e9b-f1df1e4922ec
ms.date: 06/08/2017
---


# PropertyPages Object (Outlook)

Contains the custom property pages that have been added to the Microsoft Outlook **Options** dialog box or to the folder **Properties** dialog box.


## Remarks

You receive a  **PropertyPages** object as a parameter of the **[OptionsPagesAdd](Outlook.Application.OptionsPagesAdd.md)** event. Use the **[Add](Outlook.PropertyPages.Add.md)** method to add a **[PropertyPage](Outlook.PropertyPage.md)** object to the **PropertyPages** object.


 **Note**  If more than one program handles the  **OptionsPagesAdd** event, the order in which the programs receive the event (and therefore, the order in which pages are added to the **PropertyPages** object) cannot be guaranteed.


## Methods



|**Name**|
|:-----|
|[Add](Outlook.PropertyPages.Add.md)|
|[Item](Outlook.PropertyPages.Item.md)|
|[Remove](Outlook.PropertyPages.Remove.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.PropertyPages.Application.md)|
|[Class](Outlook.PropertyPages.Class.md)|
|[Count](Outlook.PropertyPages.Count.md)|
|[Parent](Outlook.PropertyPages.Parent.md)|
|[Session](propertypages-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
