using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Mail.OutlookSignature
{
    class Program
    {
        static void AddVariable(Dictionary<string, string> variables, string variable, DirectoryEntry entry, string property)
        {
            if (entry.Properties[property] != null && entry.Properties[property].Count > 0 && entry.Properties[property].Value != null)
            {
                variables[variable] = entry.Properties[property].Value.ToString();
            }
            else
            {
                variables[variable] = "";
            }
        }

        static string ReplaceVariables(Dictionary<string, string> variables, string prefix, string text)
        {
            string result = text;

            foreach (var variable in variables)
                if (result.Contains(prefix + variable.Key + prefix))
                    result = result.Replace(prefix + variable.Key + prefix, variable.Value);

            return result;
        }

        static void Main(string[] args)
        {
            string templatePath = Properties.Settings.Default.TemplatePath;
            if (args.Length > 0)
                templatePath = args[0];

            if (string.IsNullOrWhiteSpace(templatePath))
            {
                Console.WriteLine("Missing template path to use, please set it via TemplatePath setting in .config file or pass it as argument.");
                Console.WriteLine("Usage: .exe <TemplatePath>");
                return;
            }

            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"ERROR: Specified Word template {templatePath} does not exist.");
                return;
            }

            try
            {
                if (Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software").OpenSubKey("Microsoft").OpenSubKey("Office") == null)
                {
                    Console.WriteLine("ERROR: Microsoft Office is not installed on this computer.");
                    return;
                }

                bool lockSignatureChanges = Properties.Settings.Default.LockSignature;

                if (!string.IsNullOrWhiteSpace(Properties.Settings.Default.LockSignatureOverrideGroupName))
                {
                    foreach (string groupName in DirectoryServicesUtilities.GetGroupsOfUser(Environment.UserName))
                    {
                        if (groupName == Properties.Settings.Default.LockSignatureOverrideGroupName)
                            lockSignatureChanges = false;
                    }
                }

                Dictionary<string, string> variables = GetLdapVariables();

                string signatureName = Properties.Settings.Default.SignatureName;

                string temp = Path.GetTempFileName();
                File.Copy(templatePath, temp, true);

                string vCard = GetQrCodeContent(variables);

                MessagingToolkit.QRCode.Codec.QRCodeEncoder qe = new MessagingToolkit.QRCode.Codec.QRCodeEncoder();
                System.Drawing.Bitmap qrCode = qe.Encode(vCard);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(temp, true))
                {
                    SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                    {
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveEndAndFootNotes = true,
                        RemoveFieldCodes = false,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveRsidInfo = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        ReplaceTabsWithSpaces = true,
                        RemoveBookmarks = true,
                        RemoveGoBackBookmark = true,
                    };
                    MarkupSimplifier.SimplifyMarkup(doc, settings);

                    var body = doc.MainDocumentPart.Document.Body;
                    var paras = body.Elements<Paragraph>();

                    List<OpenXmlElement> remove = new List<OpenXmlElement>();

                    List<Table> tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                    foreach (var table in tables)
                    {
                        foreach (var row in table.Elements<TableRow>())
                        {
                            foreach (var cell in row.Elements<TableCell>())
                            {
                                foreach (var paragraph in cell.Elements<Paragraph>())
                                {
                                    foreach (var run in paragraph.Elements<Run>())
                                    {
                                        foreach (var text in run.Elements<Text>())
                                        {
                                            text.Text = ReplaceVariables(variables, "%", text.Text);

                                            ReplaceLogo(variables, "%", doc, remove, text);

                                            ReplaceQr(qrCode, doc, remove, text);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (var para in paras)
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (Text text in run.Elements<Text>())
                            {
                                string parsed = ReplaceVariables(variables, "%", text.Text);
                                if (parsed.Contains(';'))
                                {
                                    string[] parts = parsed.Split(';');
                                    for (int i = 0; i < parts.Length - 1; i++)
                                    {
                                        text.InsertBeforeSelf(new Text(parts[i]));
                                        text.InsertBeforeSelf(new Break());
                                    }

                                    text.Text = parts[parts.Length - 1];
                                }
                                else
                                {
                                    text.Text = parsed;
                                }

                                ReplaceLogo(variables, "%", doc, remove, text);

                                ReplaceQr(qrCode, doc, remove, text);
                            }
                        }
                    }

                    foreach (var elem in remove)
                        elem.Remove();
                }

                var wordApp = new Microsoft.Office.Interop.Word.Application();
                string officeVersion = wordApp.Version.ToString();

                // disable roaming signatures as we are enforcing local template
                var outlookSetup = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software").
                    OpenSubKey("Microsoft").
                    OpenSubKey("Office").
                    OpenSubKey(officeVersion).
                    OpenSubKey("Outlook").
                    OpenSubKey("Setup", true);

                outlookSetup.SetValue("DisableRoamingSignaturesTemporaryToggle", 1, RegistryValueKind.DWord);

                string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures\\";

                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);


                var currentDoc = wordApp.Documents.Open(temp);
                currentDoc.SaveAs(path + signatureName + ".rtf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatRTF);
                currentDoc.SaveAs(path + signatureName + ".txt", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatEncodedText);
                currentDoc.SaveAs(path + signatureName + ".htm", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatHTML);
                currentDoc.Close();


                var mailSettings = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software").
                    OpenSubKey("Microsoft").
                    OpenSubKey("Office").
                    OpenSubKey(officeVersion).
                    OpenSubKey("Common").
                    OpenSubKey("MailSettings", true);

                // Unlock changing signature
                if (mailSettings.GetValue("NewSignature", null) != null)
                    mailSettings.DeleteValue("NewSignature");

                if (mailSettings.GetValue("ReplySignature", null) != null)
                    mailSettings.DeleteValue("ReplySignature");

                // This sets this signature as default
                wordApp.EmailOptions.EmailSignature.NewMessageSignature = signatureName;
                wordApp.EmailOptions.EmailSignature.ReplyMessageSignature = signatureName;

                wordApp.Quit();

                if (lockSignatureChanges)
                {
                    // When we write this registry keys, it blocks changing signature in outlook
                    mailSettings.SetValue("ReplySignature", signatureName);
                    mailSettings.SetValue("NewSignature", signatureName);
                }

                File.Delete(temp);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
        }

        private static string GetQrCodeContent(Dictionary<string, string> variables)
        {
            string vCard = @"BEGIN:VCARD
VERSION:2.1
N:" + variables["sn"] + ";" + variables["givenName"] + @"
FN:" + variables["displayName"] + @"
TITLE:" + variables["title"] + @"
ORG:" + variables["company"] + @"
ADR;WORK:;;" + variables["streetAddress"] + @";;" + variables["l"] + ";" + variables["postalCode"] + @";" + variables["c"] + @"
TEL;WORK:" + variables["telephoneNumber"] + @"
TEL;CELL:" + variables["mobile"] + @"
EMAIL:" + variables["mail"] + @"
END:VCARD";
            return vCard;
        }

        private static Dictionary<string, string> GetLdapVariables()
        {
            WindowsPrincipal wp = new WindowsPrincipal(WindowsIdentity.GetCurrent());
            String username = wp.Identity.Name;

            Dictionary<string, string> variables = new Dictionary<string, string>();
            var principalContext = new PrincipalContext(ContextType.Domain);
            var userPrincipal = UserPrincipal.FindByIdentity(principalContext, System.DirectoryServices.AccountManagement.IdentityType.SamAccountName, username);
            if (userPrincipal != null)
            {
                DirectoryEntry directoryEntry = userPrincipal.GetUnderlyingObject() as DirectoryEntry;

                AddVariable(variables, "givenName", directoryEntry, "givenName");
                AddVariable(variables, "sn", directoryEntry, "sn");
                AddVariable(variables, "displayName", directoryEntry, "displayName");
                AddVariable(variables, "department", directoryEntry, "department");
                AddVariable(variables, "company", directoryEntry, "company");
                AddVariable(variables, "telephoneNumber", directoryEntry, "telephoneNumber");
                AddVariable(variables, "mobile", directoryEntry, "mobile");
                AddVariable(variables, "mail", directoryEntry, "mail");
                AddVariable(variables, "physicalDeliveryOfficeName", directoryEntry, "physicalDeliveryOfficeName");
                AddVariable(variables, "postalCode", directoryEntry, "postalCode");
                AddVariable(variables, "streetAddress", directoryEntry, "streetAddress");
                AddVariable(variables, "title", directoryEntry, "title");
                AddVariable(variables, "c", directoryEntry, "c");
                AddVariable(variables, "l", directoryEntry, "l");
                AddVariable(variables, "st", directoryEntry, "st");

                // Expand country code to name
                if (string.IsNullOrEmpty(variables["c"]) == false)
                {
                    variables["country"] = "";

                    if (Countries.CountryNames.ContainsKey(variables["c"]))
                    {
                        variables["country"] = Countries.CountryNames[variables["c"]];
                    }
                }
            }
            return variables;
        }

        private static void ReplaceQr(System.Drawing.Bitmap qrCode, WordprocessingDocument doc, List<OpenXmlElement> remove, Text text)
        {
            if (text.Text.Contains("%QR%"))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (MemoryStream ms = new MemoryStream())
                {
                    qrCode.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Seek(0, SeekOrigin.Begin);
                    imagePart.FeedData(ms);
                }

                Drawing img = WordUtilities.PrepareImageDrawing(doc, mainPart.GetIdOfPart(imagePart), qrCode.Width, qrCode.Height, 1500000L);

                text.Parent.InsertAfter<Drawing>(img, text);
                remove.Add(text);
            }
        }

        private static void ReplaceLogo(Dictionary<string, string> variables, string prefix, WordprocessingDocument doc, List<OpenXmlElement> remove, Text text)
        {
            var match = System.Text.RegularExpressions.Regex.Match(text.Text, prefix + "LOGO(:([^" + prefix + "]+))?" + prefix);
            if (match.Success)
            {
                try
                {
                    string logoPath = match.Groups[2].Value;
                    logoPath = ReplaceVariables(variables, "$", logoPath);
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    using (FileStream stream = new FileStream(logoPath, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }

                    System.Drawing.Image logo = System.Drawing.Image.FromFile(logoPath);

                    Drawing img = WordUtilities.PrepareImageDrawing(doc, mainPart.GetIdOfPart(imagePart), logo.Width, logo.Height, 1500000L);

                    text.Parent.InsertAfter<Drawing>(img, text);
                }
                catch { }
                remove.Add(text);
            }
        }
    }
}
