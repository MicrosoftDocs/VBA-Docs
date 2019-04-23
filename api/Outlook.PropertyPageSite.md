---
title: PropertyPageSite object (Outlook)
keywords: vbaol11.chm384
f1_keywords:
- vbaol11.chm384
ms.prod: outlook
api_name:
- Outlook.PropertyPageSite
ms.assetid: cdec4b4c-14b3-de0a-52c8-d5af46f4644a
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyPageSite object (Outlook)

Represents the container of a custom property page.


## Remarks

Use the  **Parent** property of the ActiveX control that implements the **[PropertyPage](Outlook.PropertyPage.md)** object associated with the **PropertyPageSite** object to return the **PropertyPageSite** object. The Declarations section of the module implementing the **PropertyPage** object must contain a declaration similar to the following.


```vb
Private myPropertyPageSite As Outlook.PropertyPageSite
```

The object is then returned from the  **Parent** property.




```vb
Set myPropertyPageSite = Parent
```

Use the  **[OnStatusChange](Outlook.PropertyPageSite.OnStatusChange.md)** method to notify Microsoft Outlook that the property page has changed.


## Methods



|Name|
|:-----|
|[OnStatusChange](Outlook.PropertyPageSite.OnStatusChange.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.PropertyPageSite.Application.md)|
|[Class](Outlook.PropertyPageSite.Class.md)|
|[Parent](Outlook.PropertyPageSite.Parent.md)|
|[Session](Outlook.PropertyPageSite.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]