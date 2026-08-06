using System.Net.Http.Headers;
using System.Security.Claims;
using System.Text.Json;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;

var builder = WebApplication.CreateBuilder(args);

// MSAL.NET-backed JWT bearer validation for tokens issued by Entra ID.
// Configuration lives in the "AzureAd" section of appsettings.json.
// EnableTokenAcquisitionToCallDownstreamApi + the ClientSecret (see user secrets) turn this
// API into a confidential client, so it can exchange the caller's token for a Graph token
// on the caller's behalf (OBO) instead of trusting the claims already in the incoming token.
builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApi(builder.Configuration.GetSection("AzureAd"))
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddInMemoryTokenCaches();

builder.Services.AddAuthorization();

builder.Services.AddHttpClient("GraphApi", client =>
{
    client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
});

var spaOrigin = builder.Configuration["SpaOrigin"] ?? "http://localhost:5173";
builder.Services.AddCors(options =>
{
    options.AddPolicy("Spa", policy => policy
        .WithOrigins(spaOrigin)
        .AllowAnyHeader()
        .AllowAnyMethod());
});

var app = builder.Build();

app.UseHttpsRedirection();
app.UseCors("Spa");
app.UseAuthentication();
app.UseAuthorization();

app.MapGet("/api/profile", async (HttpContext context, ITokenAcquisition tokenAcquisition, IHttpClientFactory httpClientFactory) =>
{
    // Confirms the token was minted for the scope the SPA requested.
    context.VerifyUserHasAnyAcceptedScope("impersonate_user");

    var user = context.User;
    var scopes = user.FindFirstValue("scp")?.Split(' ', StringSplitOptions.RemoveEmptyEntries)
        ?? Array.Empty<string>();

    // Exchange the caller's access token for a Graph token via OBO, then ask Graph for the
    // display name instead of reading it off the (unverified-for-freshness) token claims.
    var graphToken = await tokenAcquisition.GetAccessTokenForUserAsync(["User.Read"]);

    using var graphClient = httpClientFactory.CreateClient("GraphApi");
    graphClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphToken);

    using var graphResponse = await graphClient.GetAsync("me?$select=displayName");
    graphResponse.EnsureSuccessStatusCode();

    using var graphPayload = JsonDocument.Parse(await graphResponse.Content.ReadAsStreamAsync());
    var displayName = graphPayload.RootElement.GetProperty("displayName").GetString();

    var profile = new
    {
        NameFromGraphAPI = displayName,
        PreferredUsername = user.FindFirstValue("preferred_username"),
        ObjectId = user.GetObjectId(),
        TenantId = user.GetTenantId(),
        Scopes = scopes,
    };

    return Results.Ok(profile);
})
.RequireAuthorization();

app.Run();
