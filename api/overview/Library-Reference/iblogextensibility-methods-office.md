---
title: IBlogExtensibility Methods (Office)
ms.prod: office
ms.assetid: 9898ea41-5954-4e58-908f-a32993ab75d8
ms.date: 06/08/2017
localization_priority: Normal
---


# IBlogExtensibility Methods (Office)

## Methods



|Name|Description|
|:-----|:-----|
|[BlogProviderProperties](../../Office.IBlogExtensibility.BlogProviderProperties.md)|Contains information about the provider.|
|[GetCategories](../../Office.IBlogExtensibility.GetCategories.md)|This method returns the list of blog categories for an account so Microsoft Word can populate the categories dropdown list.|
|[GetRecentPosts](../../Office.IBlogExtensibility.GetRecentPosts.md)|Returns the list of the user's last fifteen blog posts that Microsoft Word then displays in the  **Open Existing Post** dialog. This method does not actually return the blog post contents.|
|[GetUserBlogs](../../Office.IBlogExtensibility.GetUserBlogs.md)|Returns the list and details of user blogs associated with the specified account.|
|[Open](../../Office.IBlogExtensibility.Open.md)|Opens the blog specified by the blog ID. It is called by the  **Open Existing Post** dialog based on the item selected by the user.|
|[PublishPost](../../Office.IBlogExtensibility.PublishPost.md)|Hands off the current post so it can be published by the provider.|
|[RepublishPost](../../Office.IBlogExtensibility.RepublishPost.md)|Hands off the current post so it can be republished by the provider.|
|[SetupBlogAccount](../../Office.IBlogExtensibility.SetupBlogAccount.md)|Called from the  **Choose Account** dialog when the provider's name is chosen in the **Blog Host** dropdown or when the user requests to change a provider's account in the **Blog Accounts** dialog box.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]