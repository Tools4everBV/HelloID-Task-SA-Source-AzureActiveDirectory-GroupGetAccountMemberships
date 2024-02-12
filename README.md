# HelloID-Task-SA-Source-AzureActiveDirectory-AccountGetGroupMemberships

## Prerequisites
- [ ] This script uses the Microsoft Graph API and requires an App Registration with App permissions:
  - [ ] Read all group memberships <b><i>GroupMember.Read.All</i></b>

## Description

This code snippet executes the following tasks:

1. Define `$groupId` based on the `selectedGroup` data source input `$datasource.selectedGroup.Id`
2. Creates a token to connect to the Graph API.
3. List all user's memberships in Azure AD using the API call: [List group members](https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0&tabs=http)
4. Return a hash table for each user account using the `Write-Output` cmdlet.

> To view an example of the data source output, please refer to the JSON code pasted below and select the `Interpreted as JSON` option in HelloID

```json
{
    "Id": "7d53a91f-dd9d-41b3-94fb-143bd2fc6854"
}
```