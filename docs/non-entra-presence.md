# Presence for non-Entra and Microsoft personal accounts

This document records the current state of presence support for Teams contacts
that are not represented by a usable Microsoft Entra user object.

## What works

For ordinary Entra-backed contacts, `m365-tui` uses Microsoft Graph presence.
The contact must have a usable Entra user GUID and the signed-in account must
have the required Graph presence permission.

F7 contact diagnostics distinguishes Graph-capable identities from identities
that cannot be queried through Microsoft Graph presence.

## Microsoft personal accounts

A Teams one-to-one contact can be returned by Graph as:

```text
microsoftAccountUserConversationMember
```

For the tested personal-account contact, Graph exposed no usable Entra GUID and
no directly usable Teams MRI:

- member resource ID: opaque non-GUID
- member `userId`: opaque non-GUID
- last-message sender ID: unavailable
- cached contact ID: opaque non-GUID
- no `8:live:`/`8:orgid:`/`8:teamsvisitor:` MRI candidate

Consequently neither Graph batch presence nor
`GET /users/{id}/presence` can be used for that contact.

This is a known limitation, not a generic failure of Graph presence. F6/F7 can
still verify that `Presence.Read.All` and normal Entra presence are working.

## Display-name resolution

Microsoft personal-account roster entries can omit both `displayName` and email
even when message sender identity later contains a usable display name. In the
tested chats, `chatMessage.from.user.displayName` was present on some peer
messages and absent on others.

`m365-tui` therefore treats message sender metadata as a fallback name source.
When processing a loaded message page it scans peer messages until it finds the
first non-empty sender `displayName`; a nameless first peer message no longer
prevents a later message on the same page from teaching the application the
contact name.

F7 reports the availability and source of name metadata in current chat state,
fresh expanded-chat data, dedicated `/members`, and recent peer messages. It also
reports how many scanned peer messages contain `displayName`, but never exports
the actual name, email address, or raw identifier.

Display-name resolution is independent of presence identity. Learning a name
from message metadata does not turn an opaque personal-account ID into an Entra
GUID and therefore does not make Graph presence available.

## Experimental Teams / Skype presence path

The codebase intentionally retains an experimental, read-only path for future
use with non-Entra contacts:

```text
Teams/Skype resource token
        ->
Teams authz / Skype token
        ->
Middle Tier externalsearchv3?includeTFLUsers=true
        ->
Teams MRI such as 8:live:...
        ->
Unified Presence Service getpresence
```

This path uses the Teams/Skype resource:

```text
https://api.spaces.skype.com
resource application ID: cc15fd57-2c6c-4117-a88c-83b1d56b4bbe
```

It is not part of the normal Graph login path and is not required for ordinary
m365-tui operation.

## Current tenant/application blocker

The current `m365-tui` app registration contains Microsoft Graph as its
configured API resource, but not Microsoft Teams Services.

Two diagnostic results were observed while investigating the optional Teams
presence path:

```text
AADSTS65001
```

when trying to redeem the existing cached refresh token for the Skype/Teams
resource, and then, during an explicit interactive resource-consent probe:

```text
AADSTS650057
```

with Entra reporting that `https://api.spaces.skype.com` is not listed in the
application registration's requested permissions.

Therefore the experiment cannot proceed in the current tenant without an
administrator/application owner adding the appropriate Microsoft Teams Services
delegated API permission and granting any required consent.

This limitation does not invalidate or remove the existing Microsoft Graph
permissions and does not prevent the normal application from working.

## `teams-consent-probe`

The explicit diagnostic command:

```text
m365 teams-consent-probe
```

is intentionally isolated from normal startup.

It:

- is never run automatically;
- does not change `M365_SCOPES`;
- does not replace the primary Graph token cache;
- does not persist the experimental access token or refresh token;
- reports only sanitized OAuth/AADSTS state;
- is useful only after the relevant Teams resource has been configured for the
  app registration.

Until that permission exists, `AADSTS650057` is the expected result.

## F7 privacy rules

Contact diagnostics must not export or log:

- access tokens;
- refresh tokens;
- Skype tokens;
- raw Teams MRIs;
- raw Graph/member IDs;
- tenant IDs;
- contact email addresses or UPNs;
- raw service responses that could contain those values.

F7 may report only safe classifications and status, for example:

```text
Microsoft personal account
MRI consumer (8:live:)
opaque non-GUID
consent required (AADSTS65001)
resource missing from app registration (AADSTS650057)
```

## Future continuation

If the tenant/application permission becomes available later:

1. run `m365 teams-consent-probe`;
2. verify that the Teams/Skype resource token can be acquired;
3. use F7 on the same personal-account contact;
4. verify Teams `authz`;
5. verify Middle Tier MRI resolution;
6. verify read-only UPS presence;
7. only then consider enabling this path for normal personal-account presence.

The experimental code is deliberately kept so this investigation can continue
without reconstructing the authentication and diagnostics work from scratch.
