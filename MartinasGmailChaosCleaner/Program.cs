// See https://aka.ms/new-console-template for more information
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

Console.WriteLine("Hallo, Martina!");

try
{
    UserCredential credential;
    using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
    {
        string credPath = "token.json";
        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            GoogleClientSecrets.FromStream(stream).Secrets,
            new[] { GmailService.Scope.GmailReadonly },
            "user",
            CancellationToken.None,
            new FileDataStore(credPath, true)).Result;
    }

    var service = new GmailService(new BaseClientService.Initializer()
    {
        HttpClientInitializer = credential,
        ApplicationName = "Martinas Desktop Client for Gmail Api",
    });

    //// Retrieve emails
    //var emailListRequest = service.Users.Messages.List("me");
    //emailListRequest.MaxResults = 10;
    //var emailListResponse = emailListRequest.Execute();

    //foreach (var message in emailListResponse.Messages)
    //{
    //    Console.WriteLine(message.Payload);
    //}

    //retrieve detailed
    UsersResource.MessagesResource.ListRequest listRequest = service.Users.Messages.List("me");

    listRequest.LabelIds = "INBOX";
    listRequest.Q = "is:unread";
    listRequest.MaxResults = 100;

    ListMessagesResponse listMessageResponse = listRequest.Execute();

    if (listMessageResponse != null && listMessageResponse.Messages != null)
    {
        var allMessages = new Dictionary<string, int>();

        foreach (/*Message msg in listMessageResponse.Messages*/ var batch in listMessageResponse.Messages.Take(100))
        {
            //UsersResource.MessagesResource.GetRequest message = service.Users.Messages.Get("me", msg.Id);

            //Message msgContent = message.Execute();

            var request = service.Users.Messages.Get("me", batch.Id);
            request.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Metadata; // Get metadata only
            request.MetadataHeaders = "From"; // Only fetch the 'From' field

            Message msgContent = request.Execute();

            string fromHeader = msgContent.Payload.Headers
                .FirstOrDefault(h => h.Name == "From")?.Value ?? "Unknown";

            // Extract the email address only (remove name part)
            string email = ExtractEmail(fromHeader);

            if (msgContent != null)
            {
                if (allMessages.ContainsKey(email))
                    allMessages[email]++;
                else
                    allMessages[email] = 1;

                //var filteredFromField = msgContent.Payload.Headers.Where(m => m.Name.Equals("From")).FirstOrDefault();

                //if (filteredFromField != null)
                //{
                //    if (allMessages.ContainsKey(filteredFromField.Value))
                //    {
                //        allMessages[filteredFromField.Value]++; // Increase count if sender exists
                //    }
                //    else
                //    {
                //        allMessages[filteredFromField.Value] = 1; // Add new sender with count 1
                //    }
                //}

                //foreach (var msgParts in msgContent.Payload.Headers)
                //{
                //    if (msgParts.Name == "From")
                //    {
                //        Console.WriteLine($"From: {msgParts.Value}");
                //    }
                //}
            }
        }

        allMessages.OrderByDescending(kvp => kvp.Value);

        // Print header
        Console.WriteLine("Sender Email".PadRight(80) + "| Count");
        Console.WriteLine(new string('-', 100)); // Separator line

        // Print results
        foreach (var kvp in allMessages)
        {
            Console.WriteLine($"{kvp.Key.PadRight(80)}| {kvp.Value}");
        }
    }

    //// Searching emails
    //var query = "subject:important";
    //var searchRequest = service.Users.Messages.List("me");
    //searchRequest.Q = query;
    //var searchResponse = searchRequest.Execute();

    //foreach (var response in searchResponse.Messages)
    //{
    //    Console.WriteLine(response.Payload);
    //}
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}


static string ExtractEmail(string fromHeader)
{
    int start = fromHeader.IndexOf('<');
    int end = fromHeader.IndexOf('>');

    if (start != -1 && end != -1)
        return fromHeader.Substring(start + 1, end - start - 1); // Extract text inside < >

    return fromHeader; // If no <>, return as is
}


