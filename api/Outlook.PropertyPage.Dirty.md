---
title: PropertyPage.Dirty property (Outlook)
keywords: vbaol11.chm382
f1_keywords:
- vbaol11.chm382
ms.prod: outlook
api_name:
- Outlook.PropertyPage.Dirty
ms.assetid: fb654f40-9b80-654c-395a-811923dfb903
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyPage.Dirty property (Outlook)

Returns a **Boolean** value that indicates whether the contents of a custom property page have been altered. Read-only.


## Syntax

_expression_.**Dirty** (_Dirty_)

_expression_ A variable that represents a [PropertyPage](Outlook.PropertyPage.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Dirty_|Required| **Boolean**| **True** indicates that the contents of a custom property page has been altered.|

## Remarks

The ActiveX control that implements the  **[PropertyPage](Outlook.PropertyPage.md)** object sets the value of this property, and Microsoft Outlook queries this in response to the **[OnStatusChange](Outlook.PropertyPageSite.OnStatusChange.md)** method of a **[PropertyPageSite](Outlook.PropertyPageSite.md)** object.


## Example

This Visual Basic for Applications (VBA) example returns the value of the  **[Dirty](Outlook.PropertyPage.Dirty.md)** property as the value of a global variable.


```vb
Private Property Get PropertyPage_Dirty() As Boolean 
 PropertyPage_Dirty = globDirty 
End Property
```


## See also


[PropertyPage Object](Outlook.PropertyPage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]