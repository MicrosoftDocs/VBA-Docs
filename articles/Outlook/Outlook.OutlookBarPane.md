---
title: OutlookBarPane Object (Outlook)
keywords: vbaol11.chm3003
f1_keywords:
- vbaol11.chm3003
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane
ms.assetid: f8e6aa05-7a66-64f2-5a6a-ea639b6bbc59
ms.date: 06/08/2017
---


# OutlookBarPane Object (Outlook)

Represents the  **Shortcuts** pane in an explorer window.


## Remarks

Use the  **[Item](Outlook.Panes.Item.md)** method to retrieve the **OutlookBarPane** object from a **[Panes](Outlook.Panes.md)** object. Because the **[Name](Outlook.OutlookBarPane.Name.md)** property is the default property of the **OutlookBarPane** object, you can identify the **OutlookBarPane** object by name. For example:


## Example

The following example retrieves an  **OutlookBarPane** object by name.


```
Set myOlBarPane = myPanes.Item("OutlookBar")
```


## Events



|**Name**|
|:-----|
|[BeforeNavigate](Outlook.OutlookBarPane.BeforeNavigate.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.OutlookBarPane.Application.md)|
|[Class](Outlook.OutlookBarPane.Class.md)|
|[Contents](Outlook.OutlookBarPane.Contents.md)|
|[Name](Outlook.OutlookBarPane.Name.md)|
|[Parent](Outlook.OutlookBarPane.Parent.md)|
|[Session](Outlook.OutlookBarPane.Session.md)|
|[Visible](outlookbarpane-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
