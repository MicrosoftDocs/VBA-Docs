---
title: WebOptions object (Publisher)
keywords: vbapb10.chm8323071
f1_keywords:
- vbapb10.chm8323071
ms.prod: publisher
api_name:
- Publisher.WebOptions
ms.assetid: 15358c46-f7ca-bc37-d7ef-7d4dbfee09a4
ms.date: 06/04/2019
localization_priority: Normal
---


# WebOptions object (Publisher)

Represents the properties of a web publication, including options for saving and encoding the publication, and enabling web-safe fonts and font schemes. The **WebOptions** object is a member of the **[Application](Publisher.Application.md)** object.
 

## Remarks

The properties of the **WebOptions** object are used to specify the behavior of web publications. This means that when any of these properties are modified, newly created web publications inherit the modified properties.
 
> [!NOTE] 
> The **WebOptions** object is available from print publications and web publications. However, the properties of this object have no effect on print publications.

Use the **[WebOptions](Publisher.Application.WebOptions.md)** property of the **Application** object to return a **WebOptions** object. 
 

## Example

The following example sets an object variable equal to the Microsoft Publisher **WebOptions** object.

```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions
```


## Properties

- [AlwaysSaveInDefaultEncoding](Publisher.WebOptions.AlwaysSaveInDefaultEncoding.md)
- [Application](Publisher.WebOptions.Application.md)
- [EmailAsImg](Publisher.WebOptions.EmailAsImg.md)
- [EnableIncrementalUpload](Publisher.WebOptions.EnableIncrementalUpload.md)
- [Encoding](Publisher.WebOptions.Encoding.md)
- [OrganizeInFolder](Publisher.WebOptions.OrganizeInFolder.md)
- [Parent](Publisher.WebOptions.Parent.md)
- [RelyOnVML](Publisher.WebOptions.RelyOnVML.md)
- [ShowOnlyWebFonts](Publisher.WebOptions.ShowOnlyWebFonts.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]