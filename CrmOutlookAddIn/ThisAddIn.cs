using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Net.Http;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CrmOutlookAddIn
{
    public partial class ThisAddIn
    {
        private Application outlookApp;
        private Items inboxItems;
        private Items sentItems;

        private Ribbon1 ribbon;

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon1(this);
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            outlookApp = this.Application;

            // Monitor Inbox
            MAPIFolder inbox = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            inboxItems = inbox.Items;
            inboxItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(InboxItemAdded);

            // Monitor Sent Items
            MAPIFolder sent = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            sentItems = sent.Items;
            sentItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(SentItemAdded);
            if (Properties.Settings.Default.Init != "initialized")
            {
                Properties.Settings.Default.Init = "initialized";
                Properties.Settings.Default.RedmineUrl = Properties.Settings.Default.RedmineUrl;
                Properties.Settings.Default.RedmineApiKey = Properties.Settings.Default.RedmineApiKey;
                Properties.Settings.Default.idprefix = Properties.Settings.Default.idprefix;
                Properties.Settings.Default.ReplyDelimiter1 = Properties.Settings.Default.ReplyDelimiter1;
                Properties.Settings.Default.ReplyDelimiter2 = Properties.Settings.Default.ReplyDelimiter2;
                Properties.Settings.Default.ReplyDelimiter3 = Properties.Settings.Default.ReplyDelimiter3;
                Properties.Settings.Default.ReplyDelimiter4 = Properties.Settings.Default.ReplyDelimiter4;
                Properties.Settings.Default.UseCurlClient = Properties.Settings.Default.UseCurlClient;

                Properties.Settings.Default.Save();
            }

            // Add a TextWriterTraceListener to log to %TEMP% directory
            string tempPath = Environment.GetEnvironmentVariable("TEMP");
            if (!string.IsNullOrEmpty(tempPath))
            {
                string logFilePath = System.IO.Path.Combine(tempPath, "CrmOutlookAddIn.log");
                var listener = new TextWriterTraceListener(logFilePath);
                listener.TraceOutputOptions = TraceOptions.DateTime; // タイムスタンプを付与
                Trace.Listeners.Add(listener);
                Trace.AutoFlush = true; // Ensure logs are written immediately
                Trace.TraceInformation("Trace listener initialized. Logging to: " + logFilePath);
            }
            else
            {
                Trace.TraceError("Failed to retrieve TEMP environment variable.");
            }
        }

        private void InboxItemAdded(object Item)
        {
            if (Item is MailItem mail)
            {
                SaveMailToRedmineAsync(mail, "Received");
            }
        }

        private void SentItemAdded(object Item)
        {
            if (Item is MailItem mail)
            {
                SaveMailToRedmineAsync(mail, "Sent");
            }
        }

        private void NotifyUserOfError(System.Exception ex, string mailSubject)
        {
            // Notify error with dialog
            System.Windows.Forms.MessageBox.Show(
                $"An error occurred while saving the email with subject '{mailSubject}' to the database.\n\n" +
                $"Error Details:\n{ex.Message}\n\n" +
                "Please check the system logs for more information.",
                "Database Save Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error
            );
        }
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }


        // Class for Redmine user information
        public class RedmineUser
        {
            public int id { get; set; }
            public string mail { get; set; }
            public List<CustomField> custom_fields { get; set; }
        }

        public class CustomField
        {
            public int id { get; set; }
            public string name { get; set; }
            public string value { get; set; }
        }


        // SaveMailToRedmineAsync modification
        private async Task SaveMailToRedmineAsync(MailItem mail, string direction)
        {
            const int maxRetryCount = 5;

            for (int retryCount = 0; retryCount < maxRetryCount; retryCount++)
            {
                try
                {
                    string redmineUrl = Properties.Settings.Default.RedmineUrl;
                    string apiKey = Properties.Settings.Default.RedmineApiKey;
                    string senderEmail = GetSmtpAddress(mail.Sender);

                    // Extract id:xxxx from subject
                    string issueId = ExtractIssueIdFromSubject(mail.Subject);
                    if (string.IsNullOrEmpty(issueId))
                    {
                        Trace.TraceInformation($"No valid Redmine ticket ID found in the mail subject: {mail.Subject}");
                        return; // Skip registration
                    }

                    string sentOnString = mail.SentOn.ToString("yyyy-MM-dd HH:mm");

                    // 既存コメントの重複チェックはHttpClientで実施（curlでのGETは省略）
                    using (HttpClient client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Add("X-Redmine-API-Key", apiKey);

                        // Get existing comments
                        string getUrl = $"{redmineUrl}/issues/{issueId}.json?include=journals";
                        Trace.TraceInformation($"Sending journal request to Redmine: {getUrl}");
                        HttpResponseMessage getResponse = await client.GetAsync(getUrl);
                        if (!getResponse.IsSuccessStatusCode)
                        {
                            string errorMessage = await getResponse.Content.ReadAsStringAsync();
                            Trace.TraceInformation($"Failed to get ticket information from Redmine: {getResponse.StatusCode} - {errorMessage}");
                            return; // Skip registration
                        }

                        string issueJson = await getResponse.Content.ReadAsStringAsync();

                        // Parse JSON using Newtonsoft.Json (Json.NET) and search journals for duplicate SentOn
                        var issueObj = JObject.Parse(issueJson);
                        var journals = issueObj["issue"]?["journals"] as JArray;
                        if (journals != null)
                        {
                            foreach (var journal in journals)
                            {
                                var notes = journal["notes"]?.ToString();
                                if (!string.IsNullOrEmpty(notes) && notes.Contains($"SentOn: {sentOnString}"))
                                {
                                    Trace.TraceInformation($"A comment with the same SentOn already exists, skipping registration: {sentOnString}");
                                    return; // Avoid duplicate comments
                                }
                            }
                        }
                    }

                    string trimmedBody = TrimQuotedText(mail.Body);

                    var issueContent = new
                    {
                        issue = new
                        {
                            notes = $"SentOn: {sentOnString}\n" +
                                $"Subject: {mail.Subject ?? "No Subject"}\n" +
                                $"Sender: {senderEmail}\n" +
                                $"Recipients: {string.Join(";", mail.Recipients.Cast<Recipient>().Select(r => (string)GetSmtpAddress(r.AddressEntry)))}\n\n" +
                                $"{trimmedBody?.Substring(0, Math.Min(trimmedBody.Length, 1000)) ?? "No Body"}"
                        }
                    };

                    string jsonBody = JsonConvert.SerializeObject(issueContent);

                    if (Properties.Settings.Default.UseCurlClient)
                    {
                        // --- curlコマンドでPUT ---
                        string requestUrl = $"{redmineUrl}/issues/{issueId}.json";
                        // Windowsコマンドライン用にエスケープ
                        string escapedJson = jsonBody.Replace("\"", "\\\"").Replace("%", "%%");

                        // curlコマンド組み立て
                        string arguments = $"-X PUT \"{requestUrl}\" -H \"X-Redmine-API-Key: {apiKey}\" -H \"Content-Type: application/json\" -d \"{escapedJson}\"";

                        Trace.TraceInformation($"Executing curl: curl {arguments}");

                        var psi = new ProcessStartInfo
                        {
                            FileName = "curl",
                            Arguments = arguments,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };

                        using (var process = new Process { StartInfo = psi })
                        {
                            process.Start();
                            string output = process.StandardOutput.ReadToEnd();
                            string error = process.StandardError.ReadToEnd();
                            process.WaitForExit();

                            Trace.TraceInformation($"curl output: {output}");
                            if (process.ExitCode != 0)
                            {
                                Trace.TraceError($"curl error: {error}");
                                throw new System.Exception($"curl failed: {error}");
                            }
                        }
                    }
                    else
                    {
                        // --- 既存のHttpClientでPUT ---
                        using (HttpClient client = new HttpClient())
                        {
                            client.DefaultRequestHeaders.Add("X-Redmine-API-Key", apiKey);

                            var content = new StringContent(
                                jsonBody,
                                Encoding.UTF8,
                                "application/json"
                            );

                            string requestUrl = $"{redmineUrl}/issues/{issueId}.json";
                            Trace.TraceInformation($"Sending request to Redmine: {requestUrl}");
                            Trace.TraceInformation(jsonBody);

                            HttpResponseMessage response = await client.PutAsync(requestUrl, content);

                            if (!response.IsSuccessStatusCode)
                            {
                                string errorMessage = await response.Content.ReadAsStringAsync();
                                throw new System.Exception($"Failed to register note to Redmine: {response.StatusCode} - {errorMessage}");
                            }
                        }
                    }

                    // 成功した場合はループを抜ける
                    return;
                }
                catch (System.Exception ex)
                {
                    Trace.TraceError($"Error occurred while registering note to Redmine (Attempt {retryCount}/{maxRetryCount}): {ex.Message}\n{ex}");
                    if (retryCount == maxRetryCount - 1)
                    {
                        NotifyUserOfError(ex, mail.Subject);
                        throw; // 最後のリトライでエラーが発生した場合は通知
                    }
                    // リトライ前に少し待機
                    await Task.Delay(1000 * (retryCount + 1)); // リトライ間隔を増加させる
                }
            }
        }


        private string ExtractIssueIdFromSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return null;
            }

            string idprefix = Properties.Settings.Default.idprefix;

            // Extract ticket id using regular expression
            var match = System.Text.RegularExpressions.Regex.Match(subject, $"{idprefix}(\\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : null;
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        private string TrimQuotedText(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return body;
            }

            // Get ReplyDelimiter sequentially from app.config
            List<string> replyDelimiters = new List<string>
            {
                Properties.Settings.Default.ReplyDelimiter1,
                Properties.Settings.Default.ReplyDelimiter2,
                Properties.Settings.Default.ReplyDelimiter3,
                Properties.Settings.Default.ReplyDelimiter4
            };

            // Detect previous mail part using regular expression
            foreach (var delimiter in replyDelimiters)
            {
                var match = Regex.Match(body, delimiter, RegexOptions.Multiline | RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return body.Substring(0, match.Index).Trim(); // Return part before header
                }
            }

            return body; // Return as is if previous mail part is not found
        }

        private string GetSmtpAddress(AddressEntry addressEntry)
        {
            if (addressEntry != null)
            {
                if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    var exchUser = addressEntry.GetExchangeUser();
                    if (exchUser != null)
                    {
                        return exchUser.PrimarySmtpAddress;
                    }
                }
                else
                {
                    return addressEntry.Address;
                }
            }
            return "Unknown";
        }

        public void SaveToRedmine(IRibbonControl control)
        {
            string subject = null;
            try
            {
                string redmineUrl = Properties.Settings.Default.RedmineUrl;

                // Get the selected mail item
                var explorer = this.Application.ActiveExplorer();
                if (explorer.Selection.Count > 0 && explorer.Selection[1] is MailItem mail)
                {
                    subject = Uri.EscapeDataString(mail.Subject ?? "No Subject");

                    // Truncate the body before escaping
                    string rawBody = mail.Body ?? "No Body";
                    if (rawBody.Length > 1000)
                    {
                        rawBody = rawBody.Substring(0, 1000);
                    }

                    // Escape the truncated body
                    string description = Uri.EscapeDataString(rawBody);

                    string ticketCreationUrl = $"{redmineUrl}/issues/new?issue[subject]={subject}&issue[description]={description}";

                    Process.Start(new ProcessStartInfo
                    {
                        FileName = ticketCreationUrl,
                        UseShellExecute = true
                    });

                    Trace.TraceInformation($"Opened Redmine ticket creation page: {ticketCreationUrl}");
                }
            }
            catch (System.Exception ex)
            {
                Trace.TraceError($"Error while opening Redmine ticket creation page: {ex.Message}\n{ex}");
                NotifyUserOfError(ex, subject);
                throw;
            }
        }
    }
}
