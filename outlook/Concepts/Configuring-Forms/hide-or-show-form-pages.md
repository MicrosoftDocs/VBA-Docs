---
title: Hide or Show Form Pages
ms.prod: outlook
ms.assetid: 7efb2561-27f6-002e-8b7f-f1cffc0c8a4e
ms.date: 06/08/2017
localization_priority: Normal
---


# Hide or Show Form Pages

When you customize an Outlook form, the procedure that you use for form pages is different from the one that you use for form regions.


## Forms customized with form regions

The  [GetFormRegionStorage](../../../api/Outlook.FormRegionStartup.GetFormRegionStorage.md) method controls which form region is displayed on a form by calling the appropriate form regions storage file (.ofs). Form regions cannot be hidden or shown at run time once the form storage has been returned. For more information, see [How to: Create a Form Region](../Outlook-Forms/create-a-form-region.md) and [Extending a Form Region with an Add-in](../Specifying-Form-Behavior/extending-a-form-region-with-an-add-in.md).


## Forms customized with form pages


1. In the Forms Designer, click the page that you want to hide or show. 
    
2. On the **Developer** tab, in the **Form** group, click **Page**, and then click **Display This Page**.
    

 **Note**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]