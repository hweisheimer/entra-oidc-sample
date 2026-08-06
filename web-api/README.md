# web-api

A minimal ASP.NET Core Web API that validates Entra ID access tokens (via
[Microsoft.Identity.Web](https://github.com/AzureAD/microsoft-identity-web), which sits on
top of MSAL.NET) and returns the caller's profile. The display name is fetched from
Microsoft Graph via the on-behalf-of (OBO) flow; the rest comes from the token's claims.

## Endpoint

`GET /api/profile` — requires a bearer token with the `impersonate_user` scope.

```json
{
  "name": "Jane Doe",
  "preferredUsername": "jane@example.com",
  "objectId": "...",
  "tenantId": "...",
  "scopes": ["impersonate_user"]
}
```

## Azure setup

This sample reuses the same app registration as `web-ui`
and exposes an API from it:

1. In the app registration, go to **Expose an API** and add scope `impersonate_user`
   (this should already exist if `web-ui` is working).
2. Update `appsettings.json` `AzureAd` section with your own `TenantId` / `ClientId` if you're
   using a different app registration.
3. Under **API permissions**, add the delegated Microsoft Graph permission `User.Read` and
   grant admin consent — the API calls Graph on the caller's behalf (OBO).
4. Under **Certificates & secrets**, create a client secret. This turns the API into a
   confidential client, which is required for the OBO token exchange. Store it in user
   secrets rather than `appsettings.json`:

   ```bash
   dotnet user-secrets set "AzureAd:ClientSecret" "<your-client-secret>"
   ```

## Run

```bash
dotnet run
```

The API listens on `https://localhost:7100` (see `Properties/launchSettings.json`) and allows
CORS from `http://localhost:5173`, the Vite dev server for `web-ui`.

