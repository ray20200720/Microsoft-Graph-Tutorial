using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;
    
    public static void InitializeGraphForUserAuth(Settings settings, Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        var options = new DeviceCodeCredentialOptions
        {
            ClientId = settings.ClientId,
            TenantId = settings.TenantId,
            DeviceCodeCallback = deviceCodePrompt,
        };

        _deviceCodeCredential = new DeviceCodeCredential(options);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName" };
        });
    }

    public static Task<MessageCollectionResponse?> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            // Only messages from Inbox folder
            .MailFolders["Inbox"]
            .Messages
            .GetAsync((config) =>
            {
                // Only request specific properties
                config.QueryParameters.Select = new[] { "from", "isRead", "receivedDateTime", "subject" };
                // Get at most 25 results
                config.QueryParameters.Top = 25;
                // Sort by received time, newest first
                config.QueryParameters.Orderby = new[] { "receivedDateTime DESC" };
            });
    }

    public static async Task SendMailAsync(string subject, string body, string recipient)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Create a new message
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = BodyType.Text
            },
            ToRecipients = new List<Recipient>
        {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }
        }
        };

        // Send the message
        await _userClient.Me
            .SendMail
            .PostAsync(new SendMailPostRequestBody
            {
                Message = message
            });
    }

    // This function serves as a playground for testing Graph snippets
    // or other code
    public static Task MakeGraphCallAsync()
    {
        // INSERT YOUR CODE HERE
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            .SendMail
            .PostAsync(new SendMailPostRequestBody
            {
                
            }); 
    }

    public static async Task<Drive?> GetMeDriveAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        
        return await _userClient.Me.Drive.GetAsync();
    }

    public static async Task<List<Drive>> GetDrivesAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        var drives = (await _userClient.Me.Drives.GetAsync())?.Value;

        if (drives == null)
        {
            throw new Exception("Drives not found");
        }

        return drives;
    }

    public async Task<Drive> GetDriveById(string driveId, GraphServiceClient client)
    {
        var drive = await client.Me.Drives[$"{driveId}"].GetAsync();

        if (drive == null)
        {
            throw new Exception("Drive not found");
        }

        return drive;
    }

    public static async Task<List<DriveItem>> ListDriveRootItemsAsync(string driveId)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        var root = await _userClient.Drives[driveId].Root.GetAsync(r => r.QueryParameters.Expand = new string[] { "children" });

        if (root == null)
        {
            throw new Exception("root not found");
        }

        var items = root.Children;

        if (items == null)
        {
            throw new Exception("Items not found");
        }

        return items;
    }

    public static async Task UploadItemToFolderAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        var drive = await _userClient.Me.Drive.GetAsync();

        using (var fileStream = new FileStream("D:/file.txt", FileMode.Open))
        {
            var uploadResult = await _userClient.Drives[drive?.Id].Root
                .ItemWithPath("file.txt")
                .Content
                .PutAsync(fileStream);
        }
    }
}