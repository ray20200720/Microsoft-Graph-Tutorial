// See https://aka.ms/new-console-template for more information
Console.WriteLine(".NET Graph Tutorial\n");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

// Greet the user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. List my inbox");
    Console.WriteLine("3. Send mail");
    Console.WriteLine("4. Make a Graph call");
    Console.WriteLine("5. Get Me Drive");
    Console.WriteLine("6. Get Drives");
    Console.WriteLine("7. List Drive’s Root Items");
    Console.WriteLine("8. Upload Item To Folder");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // List emails from user's inbox
            await ListInboxAsync();
            break;
        case 3:
            // Send an email message
            await SendMailAsync();
            break;
        case 4:
            // Run any Graph code
            await MakeGraphCallAsync();
            break;
        case 5:
            await GetMeDriveAsync();
            break;
        case 6:
            await GetDrivesAsync();
            break;
        case 7:
            Console.WriteLine("Input Drive Id:");
            var driveId = Console.ReadLine();
            await ListDriveRootItemsAsync(driveId == null? string.Empty: driveId);
            break;
        case 8:
            await UploadItemToFolder();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}

void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForUserAuth(settings,
        (info, cancel) =>
        {
            // Display the device code message to
            // the user. This tells them
            // where to go to sign in and provides the
            // code to use.
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });
}

async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        var userToken = await GraphHelper.GetUserTokenAsync();
        Console.WriteLine($"User token: {userToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}

async Task ListInboxAsync()
{
    try
    {
        var messagePage = await GraphHelper.GetInboxAsync();

        if (messagePage?.Value == null)
        {
            Console.WriteLine("No results returned.");
            return;
        }

        // Output each message's details
        foreach (var message in messagePage.Value)
        {
            Console.WriteLine($"Message: {message.Subject ?? "NO SUBJECT"}");
            Console.WriteLine($"  From: {message.From?.EmailAddress?.Name}");
            Console.WriteLine($"  Status: {(message.IsRead!.Value ? "Read" : "Unread")}");
            Console.WriteLine($"  Received: {message.ReceivedDateTime?.ToLocalTime().ToString()}");
        }

        // If NextPageRequest is not null, there are more messages
        // available on the server
        // Access the next page like:
        // var nextPageRequest = new MessagesRequestBuilder(messagePage.OdataNextLink, _userClient.RequestAdapter);
        // var nextPage = await nextPageRequest.GetAsync();
        var moreAvailable = !string.IsNullOrEmpty(messagePage.OdataNextLink);

        Console.WriteLine($"\nMore messages available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user's inbox: {ex.Message}");
    }
}

async Task SendMailAsync()
{
    try
    {
        // Send mail to the signed-in user
        // Get the user for their email address
        var user = await GraphHelper.GetUserAsync();

        var userEmail = user?.Mail ?? user?.UserPrincipalName;

        if (string.IsNullOrEmpty(userEmail))
        {
            Console.WriteLine("Couldn't get your email address, canceling...");
            return;
        }

        await GraphHelper.SendMailAsync("Testing Microsoft Graph",
            "Hello world!", userEmail);

        Console.WriteLine("Mail sent.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error sending mail: {ex.Message}");
    }
}

async Task MakeGraphCallAsync()
{
    await GraphHelper.MakeGraphCallAsync();
}

async Task GetMeDriveAsync()
{
    try
    {
        var drive = await GraphHelper.GetMeDriveAsync();
        Console.WriteLine($"Driver: {drive}");
        Console.WriteLine($"Driver ID: {drive?.Id}");
        Console.WriteLine($"Driver Name: {drive?.Name}");
        Console.WriteLine($"Driver Type: {drive?.DriveType}");
        Console.WriteLine($"Driver Description: {drive?.Description}");
        Console.WriteLine($"Driver Following: {drive?.Following}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error GetMeDriveAsync: {ex.Message}");
    }
}

async Task GetDrivesAsync()
{
    try
    {
        var drives = await GraphHelper.GetDrivesAsync();
        foreach (var drive in drives)
        {
            Console.WriteLine($"Driver: {drive}");
            Console.WriteLine($"Driver ID: {drive?.Id}");
            Console.WriteLine($"Driver Name: {drive?.Name}");
            Console.WriteLine($"Driver Type: {drive?.DriveType}");
            Console.WriteLine($"Driver Description: {drive?.Description}");
            Console.WriteLine($"Driver Following: {drive?.Following}");
            Console.WriteLine("========================================");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error GetDrivesAsync: {ex.Message}");
    }
}

async Task ListDriveRootItemsAsync(string driveId)
{
    try
    {
        var driveItems = await GraphHelper.ListDriveRootItemsAsync(driveId);
        foreach(var driveItem in driveItems)
        {
            Console.WriteLine($"DriverItem ID: {driveItem.Id}");
            Console.WriteLine($"DriverItem Name: {driveItem.Name}");
            Console.WriteLine($"DriverItem Content: {driveItem.Content}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error ListDriveRootItemsAsync: {ex.Message}");
    }
}

async Task UploadItemToFolder()
{
    try
    {
        await GraphHelper.UploadItemToFolderAsync();
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error UploadItemToFolder: {ex.Message}");
    }
}