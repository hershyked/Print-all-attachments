using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PrintAllAttachments
{
    public partial class PrintAttachmentsRibbon
    {
        private void PrintAttachmentsRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Ribbon loaded
        }

        private void btnPrintAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Get the Outlook application
                Outlook.Application outlookApp = Globals.ThisAddIn.Application;
                
                // Get the active explorer (main Outlook window)
                Outlook.Explorer explorer = outlookApp.ActiveExplorer();
                
                if (explorer == null)
                {
                    MessageBox.Show("No active Outlook window found.", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Get the selected items (emails)
                Outlook.Selection selection = explorer.Selection;
                
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Please select one or more emails with attachments.", 
                        "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Track statistics
                int emailsProcessed = 0;
                int attachmentsPrinted = 0;
                List<string> errors = new List<string>();

                // Create a temporary directory for attachments
                string tempDir = Path.Combine(Path.GetTempPath(), "OutlookAttachments_" + Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);

                try
                {
                    // Iterate through selected items
                    foreach (object item in selection)
                    {
                        Outlook.MailItem mailItem = item as Outlook.MailItem;
                        
                        if (mailItem != null)
                        {
                            emailsProcessed++;
                            
                            // Check if the email has attachments
                            if (mailItem.Attachments.Count > 0)
                            {
                                // Process each attachment
                                foreach (Outlook.Attachment attachment in mailItem.Attachments)
                                {
                                    try
                                    {
                                        // Save attachment to temp directory
                                        string tempFilePath = Path.Combine(tempDir, attachment.FileName);
                                        attachment.SaveAsFile(tempFilePath);

                                        // Print the attachment
                                        bool printed = PrintFile(tempFilePath);
                                        
                                        if (printed)
                                        {
                                            attachmentsPrinted++;
                                        }
                                        else
                                        {
                                            errors.Add($"Could not print: {attachment.FileName}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        errors.Add($"Error processing {attachment.FileName}: {ex.Message}");
                                    }
                                }
                            }
                        }
                    }

                    // Show results to user
                    string message = $"Processed {emailsProcessed} email(s).\n" +
                                   $"Printed {attachmentsPrinted} attachment(s).";
                    
                    if (errors.Count > 0)
                    {
                        message += $"\n\nErrors encountered:\n{string.Join("\n", errors)}";
                    }

                    MessageBox.Show(message, "Print Attachments Complete", 
                        MessageBoxButtons.OK, 
                        errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
                finally
                {
                    // Clean up temporary directory
                    try
                    {
                        if (Directory.Exists(tempDir))
                        {
                            Directory.Delete(tempDir, true);
                        }
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Prints a file using the default application and printer
        /// </summary>
        /// <param name="filePath">Full path to the file to print</param>
        /// <returns>True if print command was sent successfully</returns>
        private bool PrintFile(string filePath)
        {
            try
            {
                // Use the ShellExecute API to print the file
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    CreateNoWindow = true,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };

                System.Diagnostics.Process process = System.Diagnostics.Process.Start(psi);
                
                // Wait a moment for the print job to be queued
                if (process != null)
                {
                    System.Threading.Thread.Sleep(1000); // Give time for print spooler
                    
                    // Try to close the process if it's still running
                    if (!process.HasExited)
                    {
                        process.CloseMainWindow();
                        process.Close();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error printing file: {ex.Message}");
                return false;
            }
        }
    }
}
