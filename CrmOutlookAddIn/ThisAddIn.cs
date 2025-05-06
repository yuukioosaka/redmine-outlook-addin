using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;
using System.Net.Http;
using Microsoft.Office.Core;

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

            // 受信トレイの監視
            MAPIFolder inbox = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            inboxItems = inbox.Items;
            inboxItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(InboxItemAdded);

            // 送信済みトレイの監視
            MAPIFolder sent = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            sentItems = sent.Items;
            sentItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(SentItemAdded);
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
            // ダイアログでエラーを通知
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


        // Redmineユーザー情報用クラス
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


        // SaveMailToRedmineAsyncの修正
        private async Task SaveMailToRedmineAsync(MailItem mail, string direction)
        {
            try
            {
                string redmineUrl = ConfigurationManager.AppSettings["RedmineUrl"];
                string apiKey = ConfigurationManager.AppSettings["RedmineApiKey"];
                string senderEmail = GetSmtpAddress(mail.Sender);

                // 件名から id:xxxx を抽出
                string issueId = ExtractIssueIdFromSubject(mail.Subject);
                if (string.IsNullOrEmpty(issueId))
                {
                    Trace.TraceInformation($"メールの件名に有効なRedmineチケットIDが見つかりませんでした: {mail.Subject}");
                    return; // 登録をスキップ
                }

                string sentOnString = mail.SentOn.ToString("yyyy-MM-dd HH:mm:ss");

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("X-Redmine-API-Key", apiKey);

                    // 既存コメントを取得
                    string getUrl = $"{redmineUrl}/issues/{issueId}.json?include=journals";
                    HttpResponseMessage getResponse = await client.GetAsync(getUrl);
                    if (!getResponse.IsSuccessStatusCode)
                    {
                        string errorMessage = await getResponse.Content.ReadAsStringAsync();
                        throw new System.Exception($"Redmineからチケット情報の取得に失敗しました: {getResponse.StatusCode} - {errorMessage}");
                    }

                    string issueJson = await getResponse.Content.ReadAsStringAsync();
                    var issueDoc = System.Text.Json.JsonDocument.Parse(issueJson);

                    // journals配列を検索してSentOnが一致するノートがあるか確認
                    if (issueDoc.RootElement.TryGetProperty("issue", out var issueElem) &&
                        issueElem.TryGetProperty("journals", out var journalsElem))
                    {
                        foreach (var journal in journalsElem.EnumerateArray())
                        {
                            if (journal.TryGetProperty("notes", out var notesElem))
                            {
                                string notes = notesElem.GetString();
                                if (!string.IsNullOrEmpty(notes) && notes.Contains($"SentOn: {sentOnString}"))
                                {
                                    Trace.TraceInformation($"同じSentOn日時のコメントが既に存在するため登録をスキップ: {sentOnString}");
                                    return; // 重複コメントを回避
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

                    var content = new StringContent(
                        System.Text.Json.JsonSerializer.Serialize(issueContent),
                        Encoding.UTF8,
                        "application/json"
                    );

                    string requestUrl = $"{redmineUrl}/issues/{issueId}.json";
                    Trace.TraceInformation($"Redmineにリクエストを送信: {requestUrl}");

                    HttpResponseMessage response = await client.PutAsync(requestUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        string errorMessage = await response.Content.ReadAsStringAsync();
                        throw new System.Exception($"Redmineへのノート登録に失敗しました: {response.StatusCode} - {errorMessage}");
                    }
                }
            }
            catch (System.Exception ex)
            {
                Trace.TraceError($"Redmineへのノート登録中にエラーが発生しました: {ex.Message}\n{ex}");
                NotifyUserOfError(ex, mail.Subject);
                throw;
            }
        }


        private string ExtractIssueIdFromSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return null;
            }

            string idprefix = ConfigurationManager.AppSettings["idprefix"];

            // 正規表現で ticketid を抽出
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

            // app.config から連番で ReplyDelimiter を取得
            List<string> replyDelimiters = new List<string>();
            for (int i = 1; i <= 9; i++) // 最大9個まで取得
            {
                string key = $"ReplyDelimiter{i}";
                string value = ConfigurationManager.AppSettings[key];
                if (!string.IsNullOrEmpty(value))
                {
                    replyDelimiters.Add(value);
                }
            }

            // 正規表現で過去のメール部分を検出
            foreach (var delimiter in replyDelimiters)
            {
                var match = System.Text.RegularExpressions.Regex.Match(body, delimiter, System.Text.RegularExpressions.RegexOptions.Multiline);
                if (match.Success)
                {
                    return body.Substring(0, match.Index).Trim(); // ヘッダー以前の部分を返す
                }
            }

            return body; // 過去のメール部分が見つからない場合はそのまま返す
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
                string redmineUrl = ConfigurationManager.AppSettings["RedmineUrl"];

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
