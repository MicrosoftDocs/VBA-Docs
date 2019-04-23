---
title: NameSpace.Categories property (Outlook)
keywords: vbaol11.chm787
f1_keywords:
- vbaol11.chm787
ms.prod: outlook
api_name:
- Outlook.NameSpace.Categories
ms.assetid: 3963afca-3a7e-38d7-1347-7e1467be3a10
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.Categories property (Outlook)

Returns or sets a  **[Categories](Outlook.Categories.md)** object that represents the set of **[Category](Outlook.Category.md)** objects that are available to the namespace. Read/write.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a '[NameSpace](Outlook.NameSpace.md)' object.


## Remarks

This property represents the Master Category List, which is the set of  **Category** objects that can be applied to Outlook items contained by the **NameSpace** object, and applies to all users of that namespace.

This property is similar to the  **[Categories](Outlook.Store.Categories.md)** property of the **[Store](Outlook.Store.md)** object. If there are multiple accounts defined in the current profile, use the **Categories** property of the store that is associated with the specific account.


## Example

The following Visual Basic for Applications (VBA) example displays a dialog box that contains the names and identifiers for each  **Category** object that is contained in the **[Categories](Outlook.NameSpace.Categories.md)** collection associated with the default **[NameSpace](Outlook.NameSpace.md)** object.


```vb
Private Sub ListCategoryIDs() 
 
 Dim objNameSpace As NameSpace 
 
 Dim objCategory As Category 
 
 Dim strOutput As String 
 
 
 
 ' Obtain a NameSpace object reference. 
 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 
 
 ' Check whether the Categories collection for the Namespace 
 
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


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]