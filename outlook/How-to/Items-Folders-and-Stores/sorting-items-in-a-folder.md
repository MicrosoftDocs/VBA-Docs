---
title: Sorting Items in a Folder
ms.prod: outlook
ms.assetid: bc3651da-cfdb-4301-4034-bb848f371e55
ms.date: 06/08/2017
localization_priority: Normal
---


# Sorting Items in a Folder

**[Items.Sort](../../../api/Outlook.Items.Sort.md)** , **[Results.Sort](../../../api/Outlook.Results.Sort.md)**, and  **[Table.Sort](../../../api/Outlook.Table.Sort.md)** allow you to sort Outlook items. This following table compares the few methods.


|| Items.Sort| Results.Sort| Table.Sort|
|:-----|:-----|:-----|:-----|
| **Object of Sorting**|Items in an  **[Items](../../../api/Outlook.Items.md)** collection based on a folder|Items in a  **[Results](../../../api/Outlook.Results.md)** collection based on a search|Items in a  **[Table](../../../api/Outlook.Table.md)** object based on a folder or search folder|
| **Applicable Properties as Sort Fields**|Explicit built-in properties with the exceptions listed in the  **Items.Sort** topic|Explicit built-in properties with the exceptions listed in the  **Results.Sort** topic|Explicit built-in properties and custom properties with the exception of binary and multi-valued properties|
| **Referencing of Properties**|<ul><li><p>By their explicit string names in the Outlook object model</p></li><li><p>Explicit built-in properties can only be referenced by their names in English and not any other localized language</p></li></ul>|<ul><li><p>By their explicit string names in the Outlook object model</p></li><li><p>Explicit built-in properties can only be referenced by their names in English and not any other localized language</p></li></ul>|<ul><li><p>By their explicit string names only; cannot reference properties by their namespaces</p></li><li><p>Explicit built-in properties can only be referenced by their names in English and not any other localized language</p></li><li><p>Custom properties can be referenced in English or a localized language</p></li></ul>|
| **Format of Sort Fields**|<ul><li><p>Enclosing square brackets ('['']') around explicit string names is optional</p></li><li><p>Property names are not case-sensitive</p></li></ul>|<ul><li><p>Enclosing square brackets ('['']') around explicit string names is optional</p></li><li><p>Property names are not case-sensitive</p></li></ul>|<ul><li><p>Enclosing square brackets ('['']') around explicit string names is optional</p></li><li><p>Property names are not case-sensitive</p></li></ul>|
| **Error Conditions**|<ul><li><p>Returns an error if property does not exist or property is not eligible for sorting.</p></li><li><p>If sort field is an empty string, no error is returned and the collection is not sorted.</p></li></ul>|<ul><li><p>Returns an error if property does not exist or property is not eligible for sorting.</p></li><li><p>If sort field is an empty string, no error is returned and the collection is not sorted.</p></li></ul>|<ul><li><p>Returns an error if property does not exist or property is not eligible for sorting.</p></li><li><p>Returns an error if sort field is an empty string</p></li></ul>|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]