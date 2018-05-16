---
title: SharedResources Object (Access)
keywords: vbaac10.chm14648
f1_keywords:
- vbaac10.chm14648
ms.prod: access
api_name:
- Access.SharedResources
ms.assetid: 45323141-e7df-1c70-efe2-926c1990d5e0
ms.date: 06/08/2017
---


# SharedResources Object (Access)

Represents the collection of shared resources in the database.


## Remarks

The SharedResources collection contains Microsoft Office themes and images that are stored once, but used throughout the database.

 For example, you may want to display your company logo on every form that you create. In earlier versions of Access, you had to import the logo into every form. In Access, you can add the logo as a shared image. Then , it will be displayed in the **Image Gallery** that is displayed when you click the **Insert Image** dropdown menu for the **Controls** group in the **Design** tab.

Use the  **[Resources](Access.CodeProject.Resources.md)** property of the **[CodeProject](Access.CodeProject.md)** object or the **[Resources](Access.CurrentProject.Resources.md)** property of the **[CurrentProject](Access.CurrentProject.md)** object to enumerate the **SharedResources** collection.

To import an image as a  **SharedResource** object, use the **[AddSharedImage](Access.CodeProject.AddSharedImage.md)** method of the **[CodeProject](Access.CodeProject.md)** object or the **[AddSharedImage](Access.CurrentProject.AddSharedImage.md)** method of the **[CurrentProject](Access.CurrentProject.md)** object.


## Properties



|**Name**|
|:-----|
|[Application](Access.SharedResources.Application.md)|
|[Count](Access.SharedResources.Count.md)|
|[Item](Access.SharedResources.Item.md)|
|[Parent](Access.SharedResources.Parent.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
