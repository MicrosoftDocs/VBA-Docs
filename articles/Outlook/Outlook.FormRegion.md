---
title: FormRegion Object (Outlook)
keywords: vbaol11.chm3018
f1_keywords:
- vbaol11.chm3018
ms.prod: outlook
api_name:
- Outlook.FormRegion
ms.assetid: 3a0b83eb-4076-9cb3-86a9-68f9e44df89f
ms.date: 06/08/2017
---


# FormRegion Object (Outlook)

Represents a form region in an Outlook form.


## Remarks

The  **FormRegion** object allows an add-in to add code behind a form region in a custom form to modify the appearance and behavior of the form region.

To obtain an instance of the  **FormRegion** object, an add-in must implement the **[FormRegionStartup](Outlook.formregionstartup.md)** interface. Outlook allocates storage for the form region, instantiates an instance of the **FormRegion** object, and returns the **FormRegion** object in the **[GetFormRegionStorage](Outlook.FormRegionStartup.GetFormRegionStorage.md)** method.

When the add-in closes the frame for the form region, the add-in must release the object for the form region.

For more infomation on programming a form region, see [Extending a Form Region with an Add-in](http://msdn.microsoft.com/library/b1a28a20-a0b8-cc57-7672-da51ec8bb097%28Office.15%29.aspx).


## Events



|**Name**|
|:-----|
|[Close](Outlook.FormRegion.Close.md)|
|[Expanded](Outlook.FormRegion.Expanded.md)|

## Methods



|**Name**|
|:-----|
|[Reflow](Outlook.FormRegion.Reflow.md)|
|[Select](Outlook.FormRegion.Select.md)|
|[SetControlItemProperty](Outlook.FormRegion.SetControlItemProperty.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.FormRegion.Application.md)|
|[Class](Outlook.FormRegion.Class.md)|
|[Detail](Outlook.FormRegion.Detail.md)|
|[DisplayName](Outlook.FormRegion.DisplayName.md)|
|[EnableAutoLayout](Outlook.FormRegion.EnableAutoLayout.md)|
|[Form](Outlook.FormRegion.Form.md)|
|[FormRegionMode](Outlook.FormRegion.FormRegionMode.md)|
|[Inspector](Outlook.FormRegion.Inspector.md)|
|[InternalName](Outlook.FormRegion.InternalName.md)|
|[IsExpanded](Outlook.FormRegion.IsExpanded.md)|
|[Item](Outlook.FormRegion.Item.md)|
|[Language](Outlook.FormRegion.Language.md)|
|[Parent](Outlook.FormRegion.Parent.md)|
|[Session](Outlook.FormRegion.Session.md)|
|[SuppressControlReplacement](Outlook.FormRegion.SuppressControlReplacement.md)|
|[Visible](Outlook.FormRegion.Visible.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
