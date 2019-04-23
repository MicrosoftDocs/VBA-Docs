---
title: IBlogExtensibility members (Office)
ms.prod: office
ms.assetid: 55f27978-9b18-f9a5-c276-298b2539ec3c
ms.date: 01/30/2019
localization_priority: Normal
---


# IBlogExtensibility members (Office)

An object that provides the ability to manipulate blog entries.


## Methods

|Name|Description|
|:-----|:-----|
|[BlogProviderProperties](../../Office.IBlogExtensibility.BlogProviderProperties.md)|Contains information about the provider.|
|[GetCategories](../../Office.IBlogExtensibility.GetCategories.md)|Returns the list of blog categories for an account so Microsoft Word can populate the categories dropdown list.|
|[GetRecentPosts](../../Office.IBlogExtensibility.GetRecentPosts.md)|Returns the list of the user's last fifteen blog posts that Microsoft Word then displays in the **Open Existing Post** dialog. This method does not actually return the blog post contents.|
|[GetUserBlogs](../../Office.IBlogExtensibility.GetUserBlogs.md)|Returns the list and details of user blogs associated with the specified account.|
|[Open](../../Office.IBlogExtensibility.Open.md)|Opens the blog specified by the blog ID. It is called by the **Open Existing Post** dialog based on the item selected by the user.|
|[PublishPost](../../Office.IBlogExtensibility.PublishPost.md)|Hands off the current post so it can be published by the provider.|
|[RepublishPost](../../Office.IBlogExtensibility.RepublishPost.md)|Hands off the current post so it can be republished by the provider.|
|[SetupBlogAccount](../../Office.IBlogExtensibility.SetupBlogAccount.md)|Called from the **Choose Account** dialog when the provider's name is chosen in the **Blog Host** dropdown or when the user requests to change a provider's account in the **Blog Accounts** dialog box.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]