---
title: PropertyPages object (Outlook)
keywords: vbaol11.chm160
f1_keywords:
- vbaol11.chm160
ms.prod: outlook
api_name:
- Outlook.PropertyPages
ms.assetid: 9850ae7b-f167-d3b2-2e9b-f1df1e4922ec
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyPages object (Outlook)

Contains the custom property pages that have been added to the Microsoft Outlook **Options** dialog box or to the folder **Properties** dialog box.


## Remarks

You receive a  **PropertyPages** object as a parameter of the **[OptionsPagesAdd](Outlook.Application.OptionsPagesAdd.md)** event. Use the **[Add](Outlook.PropertyPages.Add.md)** method to add a **[PropertyPage](Outlook.PropertyPage.md)** object to the **PropertyPages** object.


> [!NOTE] 
> If more than one program handles the  **OptionsPagesAdd** event, the order in which the programs receive the event (and therefore, the order in which pages are added to the **PropertyPages** object) cannot be guaranteed.


## Methods



|Name|
|:-----|
|[Add](Outlook.PropertyPages.Add.md)|
|[Item](Outlook.PropertyPages.Item.md)|
|[Remove](Outlook.PropertyPages.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.PropertyPages.Application.md)|
|[Class](Outlook.PropertyPages.Class.md)|
|[Count](Outlook.PropertyPages.Count.md)|
|[Parent](Outlook.PropertyPages.Parent.md)|
|[Session](Outlook.PropertyPages.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]