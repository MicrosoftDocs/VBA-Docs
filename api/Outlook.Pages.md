---
title: Pages object (Outlook)
keywords: vbaol11.chm390
f1_keywords:
- vbaol11.chm390
ms.prod: outlook
api_name:
- Outlook.Pages
ms.assetid: ed4dd77e-b339-7f43-d036-c02daa69d5b8
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages object (Outlook)

Contains pages that represent the pages of an Inspector window.


## Remarks

Every  **[Inspector](Outlook.Inspector.md)** object has a **Pages** object defined, which is empty (count 0) if the Outlook item has never been customized before.

Use the  **[ModifiedFormPages](Outlook.Inspector.ModifiedFormPages.md)** property to return the **Pages** object from an **Inspector** object.

Use the  **[Add](Outlook.Pages.Add.md)** method to create a custom page (you can add as many as 5 customizable pages). Use the ** _Name_** argument of the **Add** method to set the display name of the returned page. In addition to adding custom pages, you can use the _Name_ argument to return the main page of an **Inspector** object for modification.

Use  **ModifiedFormPages** (_index_), where _index_ is the name or index number, to return a single page from a **Pages** object.


## Example



The following example returns the  **Pages** object for the active **Inspector**.




```vb
Set myPages = myItem.GetInspector.ModifiedFormPages
```

The following example returns a custom page with a default name (such as "Custom1").




```vb
Set myPage = myPages.Add
```

The following example returns a custom page named "My Page."






```vb
Set myPage = myPages.Add("My Page")
```

The following example returns the Message page if the Inspector contains a mail message.




```vb
Set myPage = myPages.Add("Message")
```

The following example returns the General (main) page if the inspector contains a contact.




```vb
Set myPage = myPages.Add("General")
```


## Methods



|Name|
|:-----|
|[Add](Outlook.Pages.Add.md)|
|[Item](Outlook.Pages.Item.md)|
|[Remove](Outlook.Pages.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Pages.Application.md)|
|[Class](Outlook.Pages.Class.md)|
|[Count](Outlook.Pages.Count.md)|
|[Parent](Outlook.Pages.Parent.md)|
|[Session](Outlook.Pages.Session.md)|

## See also

- [Object model (Outlook)](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]