---
title: OutlookBarStorage Object (Outlook)
keywords: vbaol11.chm367
f1_keywords:
- vbaol11.chm367
ms.prod: outlook
api_name:
- Outlook.OutlookBarStorage
ms.assetid: e6dc8dc0-bae4-f59b-c991-1421b280de38
ms.date: 06/08/2017
---


# OutlookBarStorage Object (Outlook)

Represents the storage for objects in the Outlook Bar.


## Remarks

Use the  **[Contents](Outlook.OutlookBarPane.Contents.md)** property of an **[OutlookBarPane](Outlook.OutlookBarPane.md)** object to retrieve the **OutlookBarStorage** object for the Outlook Bar.

Use the  **[Groups](Outlook.OutlookBarStorage.Groups.md)** property to retrieve the **[OutlookBarGroups](Outlook.OutlookBarGroups.md)** object for the Outlook Bar.


## Example

The following example retrieves an  **OutlookBarStorage** object by name.


```
Set myOLBarStorage = myPanes.Item("OutlookBar").Contents
```


## Properties



|**Name**|
|:-----|
|[Application](Outlook.OutlookBarStorage.Application.md)|
|[Class](Outlook.OutlookBarStorage.Class.md)|
|[Groups](Outlook.OutlookBarStorage.Groups.md)|
|[Parent](Outlook.OutlookBarStorage.Parent.md)|
|[Session](outlookbarstorage-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
