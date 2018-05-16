---
title: AccountSelector Object (Outlook)
keywords: vbaol11.chm3456
f1_keywords:
- vbaol11.chm3456
ms.prod: outlook
api_name:
- Outlook.AccountSelector
ms.assetid: 846f176e-5680-a214-7624-75f3a524c989
ms.date: 06/08/2017
---


# AccountSelector Object (Outlook)

Provides the ability to obtain the account that is selected in the Microsoft Office Backstage view for the parent  **[Explorer](Outlook.Explorer.md)** object.


## Remarks

The  **AccountSelector** object has the **Explorer** object as its parent object. You can obtain an instance of the **AccountSelector** object from the **[AccountSelector](Outlook.Explorer.AccountSelector.md)** property of the **Explorer** object.

The  **AccountSelector** object provides a **[SelectedAccount](Outlook.AccountSelector.SelectedAccount.md)** property that returns the current account that has been selected in the Backstage view. The object also provides a **[SelectedAccountChange](Outlook.AccountSelector.SelectedAccountChange.md)** event that fires when the user has changed the account in the Backstage view.


## Events



|**Name**|
|:-----|
|[SelectedAccountChange](Outlook.AccountSelector.SelectedAccountChange.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.AccountSelector.Application.md)|
|[Class](Outlook.AccountSelector.Class.md)|
|[Parent](Outlook.AccountSelector.Parent.md)|
|[SelectedAccount](Outlook.AccountSelector.SelectedAccount.md)|
|[Session](accountselector-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
