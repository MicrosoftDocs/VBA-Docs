---
title: Categories object (Outlook)
keywords: vbaol11.chm3178
f1_keywords:
- vbaol11.chm3178
ms.prod: outlook
api_name:
- Outlook.Categories
ms.assetid: 319efa26-269d-9f2f-c8ec-33082e80a9e2
ms.date: 06/08/2017
localization_priority: Normal
---

# Categories object (Outlook)

Represents the collection of **[Category](Outlook.Category.md)** objects that define the Master Category List for a namespace.

## Remarks

Microsoft Outlook provides a categorization system by which Outlook items can be easily identified and grouped into user-defined categories. The **Categories** object represents the set of user-defined categories available to the user of a given mailbox.

Use the **[Categories](Outlook.NameSpace.Categories.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object to obtain a **Categories** object reference, representing the Master Category List for that namespace.

Use the **[Add](Outlook.Categories.Add.md)** method to create a new **Category** object and append it to the collection. Use the **[Item](Outlook.Categories.Item.md)** method to obtain a **Category** object reference for an existing category, and the **[Remove](Outlook.Categories.Remove.md)** method to remove a **Category** object from the collection. Use the **[Count](Outlook.Categories.Count.md)** property to return the number of categories contained in the collection.

## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing the names and identifiers for each **Category** object contained in the **Categories** collection associated with the default **[NameSpace](Outlook.NameSpace.md)** object.


```vb
Private Sub ListCategoryIDs() 
 Dim objNameSpace As NameSpace 
 Dim objCategory As Category 
 Dim strOutput As String 
 
 ' Obtain a NameSpace object reference. 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 ' Check if the Categories collection for the Namespace 
 ' contains one or more Category objects. 
 If objNameSpace.Categories.Count > 0 Then 
 
 ' Enumerate the Categories collection. 
 For Each objCategory In objNameSpace.Categories 
 
 ' Add the name and ID of the Category object to 
 ' the output string. 
 strOutput = strOutput & objCategory.Name & _ 
 ": " & objCategory.CategoryID & vbCrLf 
 Next 
 End If 
 
 ' Display the output string. 
 MsgBox strOutput 
 
 ' Clean up. 
 Set objCategory = Nothing 
 Set objNameSpace = Nothing 
 
End Sub 

```


## Methods

|Name|
|:-----|
|[Add](Outlook.Categories.Add.md)|
|[Item](Outlook.Categories.Item.md)|
|[Remove](Outlook.Categories.Remove.md)|

## Properties

|Name|
|:-----|
|[Application](Outlook.Categories.Application.md)|
|[Class](Outlook.Categories.Class.md)|
|[Count](Outlook.Categories.Count.md)|
|[Parent](Outlook.Categories.Parent.md)|
|[Session](Outlook.Categories.Session.md)|

## See also

- [Categories Object Members](overview/Outlook.md)
- [Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
