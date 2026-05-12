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
using Microsoft.Win32;

namespace CrmOutlookAddIn
{
    /// <summary>
    /// レジストリから設定を管理する静的クラス。
    /// HKEY_CURRENT_USER\Software\CrmOutlookAddIn に設定を保存します。
    /// </summary>
    public static class SettingsManager
    {
        private const string RegistryKeyPath = @"Software\CrmOutlookAddIn";

        /// <summary>
        /// レジストリから値を取得します。キーが存在しない場合はデフォルト値を返します。
        /// </summary>
        private static object GetValue(string key, object defaultValue)
        {
            using (RegistryKey regKey = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                return regKey?.GetValue(key, defaultValue) ?? defaultValue;
            }
        }

        /// <summary>
        /// レジストリに値を設定します。
        /// </summary>
        private static void SetValue(string key, object value)
        {
            using (RegistryKey regKey = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                regKey?.SetValue(key, value);
            }
        }

        public static string Init
        {
            get => GetValue("Init", "").ToString();
            set => SetValue("Init", value);
        }

        public static string RedmineUrl
        {
            get => GetValue("RedmineUrl", "http://redmine.example.com").ToString();
            set => SetValue("RedmineUrl", value);
        }

        public static string RedmineApiKey
        {
            get => GetValue("RedmineApiKey", "").ToString();
            set => SetValue("RedmineApiKey", value);
        }

        public static string RedmineProjectId
        {
            get => GetValue("RedmineProjectId", "").ToString();
            set => SetValue("RedmineProjectId", value);
        }

        public static int RedmineUserId
        {
            get => Convert.ToInt32(GetValue("RedmineUserId", 0));
            set => SetValue("RedmineUserId", value);
        }

        public static string idprefix
        {
            get => GetValue("idprefix", "\\[id-").ToString();
            set => SetValue("idprefix", value);
        }

        public static string ReplyDelimiter1
        {
            get => GetValue("ReplyDelimiter1", "From:").ToString();
            set => SetValue("ReplyDelimiter1", value);
        }

        public static string ReplyDelimiter2
        {
            get => GetValue("ReplyDelimiter2", "差出人:").ToString();
            set => SetValue("ReplyDelimiter2", value);
        }

        public static string ReplyDelimiter3
        {
            get => GetValue("ReplyDelimiter3", "-----Original Message-----").ToString();
            set => SetValue("ReplyDelimiter3", value);
        }

        public static string ReplyDelimiter4
        {
            get => GetValue("ReplyDelimiter4", "From ").ToString();
            set => SetValue("ReplyDelimiter4", value);
        }

        public static bool UseCurlClient
        {
            get => Convert.ToBoolean(GetValue("UseCurlClient", false));
            set => SetValue("UseCurlClient", value);
        }
    }

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

            // 受信トレイの監視設定
            MAPIFolder inbox = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            inboxItems = inbox.Items;
            inboxItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(InboxItemAdded);

            // 送信済みアイテムの監視設定
            MAPIFolder sent = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            sentItems = sent.Items;
            sentItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(SentItemAdded);

            outlookApp.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(OutlookApp_ItemSend);

            Task.Run(() => InitializeInBackground());
        }

        /// <summary>
        /// バックグラウンドスレッドで実行する初期化処理をまとめたメソッド。
        /// </summary>
        private void InitializeInBackground()
        {
            try
            {
                if (SettingsManager.Init != "initialized")
                {
                    SettingsManager.Init = "initialized";
                    SettingsManager.idprefix = SettingsManager.idprefix;
                    SettingsManager.RedmineApiKey = SettingsManager.RedmineApiKey;
                    SettingsManager.RedmineUrl = SettingsManager.RedmineUrl;
                    SettingsManager.RedmineProjectId = SettingsManager.RedmineProjectId;
                    SettingsManager.ReplyDelimiter1 = SettingsManager.ReplyDelimiter1;
                    SettingsManager.ReplyDelimiter2 = SettingsManager.ReplyDelimiter2;
                    SettingsManager.ReplyDelimiter3 = SettingsManager.ReplyDelimiter3;
                    SettingsManager.ReplyDelimiter4 = SettingsManager.ReplyDelimiter4;
                    SettingsManager.UseCurlClient = SettingsManager.UseCurlClient;

                    Trace.TraceInformation("First time initialization. Settings will use default values from registry.");
                }

                // トレースリスナーの初期化
                string tempPath = Environment.GetEnvironmentVariable("TEMP");
                if (!string.IsNullOrEmpty(tempPath))
                {
                    string logFilePath = System.IO.Path.Combine(tempPath, "CrmOutlookAddIn.log");
                    var listener = new TextWriterTraceListener(logFilePath);
                    listener.TraceOutputOptions = TraceOptions.DateTime;
                    Trace.Listeners.Add(listener);
                    Trace.AutoFlush = true;
                    Trace.TraceInformation("Background initialization complete. Logging to: " + logFilePath);
                }
                else
                {
                    Trace.TraceError("Failed to retrieve TEMP environment variable during background initialization.");
                }
            }
            catch (System.Exception ex)
            {
                Trace.TraceError($"An error occurred during background initialization: {ex.ToString()}");
            }
        }

        /// <summary>
        /// メールの送信ボタンを押した際の処理
        /// 件名にIDが含まれていない場合、新規チケットを作成するか確認し、件名を書き換えます。
        /// </summary>
        private void OutlookApp_ItemSend(object Item, ref bool Cancel)
        {
            if (Item is MailItem mail)
            {
                try
                {
                    string projectId = SettingsManager.RedmineProjectId;
                    if (string.IsNullOrEmpty(projectId))
                    {
                        Trace.TraceInformation("RedmineProjectId is not set. Skipping automatic ticket creation.");
                        return;
                    }

                    string issueId = ExtractIssueIdFromSubject(mail.Subject);

                    // 件名にチケットIDが含まれていない場合
                    if (string.IsNullOrEmpty(issueId))
                    {
                        var dialogResult = System.Windows.Forms.MessageBox.Show(
                            "このメールの件名にはRedmineのチケットIDが含まれていません。\n新規チケットを作成してから送信しますか？\n\n" +
                            "[はい] : 新規チケットを作成して件名に追記し、送信する\n" +
                            "[いいえ] : 新規チケットは作成せず、そのまま送信する\n" +
                            "[キャンセル] : メールの送信処理自体を中止する",
                            "Redmine 新規チケット作成の確認",
                            System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                            System.Windows.Forms.MessageBoxIcon.Question
                        );

                        if (dialogResult == System.Windows.Forms.DialogResult.Cancel)
                        {
                            Cancel = true;
                            return;
                        }
                        else if (dialogResult == System.Windows.Forms.DialogResult.No)
                        {
                            return;
                        }

                        string subject = mail.Subject ?? "No Subject";
                        string senderAddress = GetSmtpAddress(mail.Sender);
                        string trimmedBody = TrimQuotedText(mail.Body);

                        int newIssueId = CreateRedmineTicketSync(subject, senderAddress, trimmedBody);

                        if (newIssueId > 0)
                        {
                            string prefix = SettingsManager.idprefix.Replace("\\", "");
                            string suffix = prefix.StartsWith("[") ? "]" : "";

                            string idString = $"{prefix}{newIssueId}{suffix}";

                            mail.Subject = $"{mail.Subject} {idString}";
                            mail.Save();

                            string url = $"{SettingsManager.RedmineUrl}/issues/{newIssueId}";
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = url,
                                UseShellExecute = true
                            });
                            Trace.TraceInformation($"Opened newly created Redmine ticket page: {url}");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Trace.TraceError($"Error in ItemSend (Ticket Creation): {ex.ToString()}");

                    var result = System.Windows.Forms.MessageBox.Show(
                        $"Redmineチケットの自動作成に失敗しました。\nこのままメールを送信しますか？\n\nエラー詳細:\n{ex.Message}",
                        "Redmine Ticket Error",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    if (result == System.Windows.Forms.DialogResult.No)
                    {
                        Cancel = true;
                    }
                }
            }
        }

        /// <summary>
        /// APIキーから自分自身のRedmineユーザーIDを取得します。
        /// </summary>
        private int GetCurrentUserIdSync()
        {
            try
            {
                int cachedId = SettingsManager.RedmineUserId;
                if (cachedId > 0) return cachedId;

                string requestUrl = $"{SettingsManager.RedmineUrl}/users/current.json";
                string apiKey = SettingsManager.RedmineApiKey;

                if (SettingsManager.UseCurlClient)
                {
                    // curl側は変更なし
                    string arguments = $"-s -X GET \"{requestUrl}\" -H \"X-Redmine-API-Key: {apiKey}\"";
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
                        process.WaitForExit();

                        if (process.ExitCode == 0)
                        {
                            var responseObj = JObject.Parse(output);
                            if (responseObj["user"] != null && responseObj["user"]["id"] != null)
                            {
                                int userId = responseObj["user"]["id"].Value<int>();
                                SettingsManager.RedmineUserId = userId;
                                return userId;
                            }
                        }
                    }
                }
                else
                {
                    // ★ HttpClient → HttpWebRequest に変更
                    var request = (HttpWebRequest)WebRequest.Create(requestUrl);
                    request.Headers.Add("X-Redmine-API-Key", apiKey);
                    request.Method = "GET";

                    using (var webResponse = (HttpWebResponse)request.GetResponse())
                    using (var reader = new System.IO.StreamReader(webResponse.GetResponseStream()))
                    {
                        string responseString = reader.ReadToEnd();
                        var responseObj = JObject.Parse(responseString);
                        if (responseObj["user"] != null && responseObj["user"]["id"] != null)
                        {
                            int userId = responseObj["user"]["id"].Value<int>();
                            SettingsManager.RedmineUserId = userId;
                            return userId;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Trace.TraceWarning($"Failed to get current user ID: {ex.Message}");
            }
            return 0;
        }

        private int CreateRedmineTicketSync(string subject, string senderAddress, string body)
        {
            string redmineUrl = SettingsManager.RedmineUrl;
            string apiKey = SettingsManager.RedmineApiKey;
            string projectId = SettingsManager.RedmineProjectId;

            if (string.IsNullOrEmpty(projectId))
            {
                throw new System.Exception("RedmineのプロジェクトID(RedmineProjectId)が設定されていません。");
            }

            string description = $"Sender: {senderAddress}\n\n" +
                                 $"{body?.Substring(0, Math.Min(body.Length, 1000)) ?? "No Body"}";

            int currentUserId = GetCurrentUserIdSync();

            var issueData = new Dictionary<string, object>
            {
                { "project_id", projectId },
                { "subject", subject },
                { "description", description }
            };

            if (currentUserId > 0)
            {
                issueData.Add("assigned_to_id", currentUserId);
            }

            var issueContent = new { issue = issueData };
            string jsonBody = JsonConvert.SerializeObject(issueContent);
            string requestUrl = $"{redmineUrl}/issues.json";

            if (SettingsManager.UseCurlClient)
            {
                // curl側は変更なし
                string escapedJson = jsonBody.Replace("\"", "\\\"").Replace("%", "%%");
                string arguments = $"-s -X POST \"{requestUrl}\" -H \"X-Redmine-API-Key: {apiKey}\" -H \"Content-Type: application/json\" -d \"{escapedJson}\"";

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

                    if (process.ExitCode != 0) throw new System.Exception($"curl failed: {error}");

                    var responseObj = JObject.Parse(output);
                    if (responseObj["issue"] != null && responseObj["issue"]["id"] != null)
                    {
                        return responseObj["issue"]["id"].Value<int>();
                    }
                    throw new System.Exception($"Failed to parse issue id from curl output. Output: {output}");
                }
            }
            else
            {
                // ★ HttpClient → HttpWebRequest に変更
                byte[] bodyBytes = Encoding.UTF8.GetBytes(jsonBody);

                var request = (HttpWebRequest)WebRequest.Create(requestUrl);
                request.Headers.Add("X-Redmine-API-Key", apiKey);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.ContentLength = bodyBytes.Length;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(bodyBytes, 0, bodyBytes.Length);
                }

                try
                {
                    using (var webResponse = (HttpWebResponse)request.GetResponse())
                    using (var reader = new System.IO.StreamReader(webResponse.GetResponseStream()))
                    {
                        string responseString = reader.ReadToEnd();
                        var responseObj = JObject.Parse(responseString);
                        if (responseObj["issue"] != null && responseObj["issue"]["id"] != null)
                        {
                            return responseObj["issue"]["id"].Value<int>();
                        }
                        throw new System.Exception($"Failed to parse issue id from API response. Response: {responseString}");
                    }
                }
                catch (WebException webEx)
                {
                    // ★ HttpWebRequest はエラーレスポンスを例外で返すため個別にハンドリング
                    if (webEx.Response is HttpWebResponse errorResponse)
                    {
                        using (var reader = new System.IO.StreamReader(errorResponse.GetResponseStream()))
                        {
                            string errorBody = reader.ReadToEnd();
                            throw new System.Exception($"Failed to create ticket: {errorResponse.StatusCode} - {errorBody}");
                        }
                    }
                    throw;
                }
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
        }

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

        private async Task SaveMailToRedmineAsync(MailItem mail, string direction)
        {
            string subject = mail.Subject;
            string senderEmail = GetSmtpAddress(mail.Sender);
            string sentOnString = mail.SentOn.ToString("yyyy-MM-dd HH:mm");
            string trimmedBody = TrimQuotedText(mail.Body);
            string recipients = string.Join(";", mail.Recipients.Cast<Recipient>().Select(r => (string)GetSmtpAddress(r.AddressEntry)));

            const int maxRetryCount = 5;

            for (int retryCount = 0; retryCount < maxRetryCount; retryCount++)
            {
                try
                {
                    string redmineUrl = SettingsManager.RedmineUrl;
                    string apiKey = SettingsManager.RedmineApiKey;


                    string issueId = ExtractIssueIdFromSubject(subject);
                    if (string.IsNullOrEmpty(issueId))
                    {
                        Trace.TraceInformation($"No valid Redmine ticket ID found in the mail subject: {subject}");
                        return;
                    }

                    using (HttpClient client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Add("X-Redmine-API-Key", apiKey);

                        string getUrl = $"{redmineUrl}/issues/{issueId}.json?include=journals";
                        Trace.TraceInformation($"Sending journal request to Redmine: {getUrl}");
                        HttpResponseMessage getResponse = await client.GetAsync(getUrl);
                        if (!getResponse.IsSuccessStatusCode)
                        {
                            string errorMessage = await getResponse.Content.ReadAsStringAsync();
                            Trace.TraceInformation($"Failed to get ticket information from Redmine: {getResponse.StatusCode} - {errorMessage}");
                            return;
                        }

                        string issueJson = await getResponse.Content.ReadAsStringAsync();

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
                                    return;
                                }
                            }
                        }
                    }

                    var issueContent = new
                    {
                        issue = new
                        {
                            notes = $"SentOn: {sentOnString}\n" +
                                $"Subject: {subject ?? "No Subject"}\n" +
                                $"Sender: {senderEmail}\n" +
                                $"Recipients: {recipients}\n\n" +
                                $"{trimmedBody?.Substring(0, Math.Min(trimmedBody.Length, 1000)) ?? "No Body"}"
                        }
                    };

                    string jsonBody = JsonConvert.SerializeObject(issueContent);

                    if (SettingsManager.UseCurlClient)
                    {
                        string requestUrl = $"{redmineUrl}/issues/{issueId}.json";
                        string escapedJson = jsonBody.Replace("\"", "\\\"").Replace("%", "%%");

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

                    return;
                }
                catch (System.Exception ex)
                {
                    Trace.TraceError($"Error occurred while registering note to Redmine (Attempt {retryCount}/{maxRetryCount}): {ex.Message}\n{ex}");
                    if (retryCount == maxRetryCount - 1)
                    {
                        NotifyUserOfError(ex, subject);
                        throw;
                    }
                    await Task.Delay(1000 * (retryCount + 1));
                }
            }
        }

        private string ExtractIssueIdFromSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return null;
            }

            string idprefix = SettingsManager.idprefix;

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

            List<string> replyDelimiters = new List<string>
            {
                SettingsManager.ReplyDelimiter1,
                SettingsManager.ReplyDelimiter2,
                SettingsManager.ReplyDelimiter3,
                SettingsManager.ReplyDelimiter4
            };

            foreach (var delimiter in replyDelimiters)
            {
                if (!string.IsNullOrEmpty(delimiter))
                {
                    var match = Regex.Match(body, delimiter, RegexOptions.Multiline | RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        return body.Substring(0, match.Index).Trim();
                    }
                }
            }

            return body;
        }

        private string GetSmtpAddress(AddressEntry addressEntry)
        {
            if (addressEntry != null)
            {
                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    try
                    {
                        return addressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS).ToString();
                    }
                    catch (System.Exception)
                    {
                        var exchUser = addressEntry.GetExchangeUser();
                        if (exchUser != null)
                        {
                            return exchUser.PrimarySmtpAddress;
                        }
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
                string redmineUrl = SettingsManager.RedmineUrl;

                var explorer = this.Application.ActiveExplorer();
                if (explorer.Selection.Count > 0 && explorer.Selection[1] is MailItem mail)
                {
                    subject = Uri.EscapeDataString(mail.Subject ?? "No Subject");

                    string rawBody = mail.Body ?? "No Body";
                    if (rawBody.Length > 1000)
                    {
                        rawBody = rawBody.Substring(0, 1000);
                    }

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
