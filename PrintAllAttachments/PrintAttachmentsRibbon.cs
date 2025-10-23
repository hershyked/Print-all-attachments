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
                    MessageBox.Show("No active Outlook window found.\n\nPlease make sure Outlook is open and try again.", 
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Get the selected items (emails)
                Outlook.Selection selection = explorer.Selection;
                
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Please select one or more emails with attachments.\n\n" +
                        "To select emails:\n" +
                        "• Click on an email to select it\n" +
                        "• Hold Ctrl and click to select multiple emails\n" +
                        "• Hold Shift and click to select a range", 
                        "No Emails Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Count total attachments to print
                int totalAttachments = 0;
                int emailsWithAttachments = 0;
                foreach (object item in selection)
                {
                    Outlook.MailItem mailItem = item as Outlook.MailItem;
                    if (mailItem != null && mailItem.Attachments.Count > 0)
                    {
                        totalAttachments += mailItem.Attachments.Count;
                        emailsWithAttachments++;
                    }
                }

                if (totalAttachments == 0)
                {
                    MessageBox.Show($"The selected email(s) do not contain any attachments.\n\n" +
                        "Please select emails that have attachments (look for the paperclip icon 📎).", 
                        "No Attachments Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Confirm before processing large batches
                if (totalAttachments > 20)
                {
                    DialogResult result = MessageBox.Show(
                        $"You are about to print {totalAttachments} attachment(s) from {emailsWithAttachments} email(s).\n\n" +
                        "This may take several minutes. Do you want to continue?",
                        "Confirm Print",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    
                    if (result != DialogResult.Yes)
                    {
                        return;
                    }
                }

                // Track statistics
                int emailsProcessed = 0;
                int attachmentsPrinted = 0;
                int attachmentsSkipped = 0;
                List<string> errors = new List<string>();

                // Create a temporary directory for attachments
                string tempDir = Path.Combine(Path.GetTempPath(), "OutlookAttachments_" + Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);

                try
                {
                    // Show initial progress message for large batches
                    if (totalAttachments > 10)
                    {
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    }

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
                                        // Check for embedded message attachments which can't be printed
                                        if (attachment.Type == (int)Outlook.OlAttachmentType.olEmbeddeditem)
                                        {
                                            attachmentsSkipped++;
                                            continue;
                                        }

                                        // Save attachment to temp directory
                                        string tempFilePath = Path.Combine(tempDir, SanitizeFileName(attachment.FileName));
                                        attachment.SaveAsFile(tempFilePath);

                                        // Print the attachment
                                        PrintResult result = PrintFile(tempFilePath, attachment.FileName);
                                        
                                        if (result.Success)
                                        {
                                            attachmentsPrinted++;
                                        }
                                        else
                                        {
                                            if (result.IsUnsupportedType)
                                            {
                                                attachmentsSkipped++;
                                            }
                                            else
                                            {
                                                errors.Add($"Could not print '{attachment.FileName}': {result.ErrorMessage}");
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        errors.Add($"Error processing '{attachment.FileName}': {ex.Message}");
                                    }
                                }
                            }
                        }
                    }

                    // Reset cursor
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

                    // Show results to user
                    string message = $"✓ Processed {emailsProcessed} email(s)\n" +
                                   $"✓ Printed {attachmentsPrinted} attachment(s) successfully";
                    
                    if (attachmentsSkipped > 0)
                    {
                        message += $"\n⚠ Skipped {attachmentsSkipped} unsupported attachment(s)";
                    }
                    
                    if (errors.Count > 0)
                    {
                        message += $"\n\n❌ Errors ({errors.Count}):\n";
                        // Limit error display to first 5 to avoid huge dialogs
                        int errorCount = Math.Min(5, errors.Count);
                        for (int i = 0; i < errorCount; i++)
                        {
                            message += $"• {errors[i]}\n";
                        }
                        if (errors.Count > 5)
                        {
                            message += $"... and {errors.Count - 5} more error(s)";
                        }
                    }

                    if (attachmentsPrinted == 0 && errors.Count == 0 && attachmentsSkipped > 0)
                    {
                        message += "\n\nℹ All attachments were skipped because they are unsupported file types.\n" +
                                  "Common unsupported types: embedded emails, calendar items.";
                    }

                    MessageBox.Show(message, "Print Attachments Complete", 
                        MessageBoxButtons.OK, 
                        errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
                finally
                {
                    // Reset cursor
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

                    // Clean up temporary directory with retry logic
                    CleanupTempDirectory(tempDir);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                MessageBox.Show($"An unexpected error occurred:\n\n{ex.Message}\n\n" +
                    "Please try again. If the problem persists, try:\n" +
                    "• Restart Outlook\n" +
                    "• Check your printer settings\n" +
                    "• Verify you have write access to the temp folder", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Result of a print operation
        /// </summary>
        private class PrintResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public bool IsUnsupportedType { get; set; }
        }

        /// <summary>
        /// Prints a file using the default application and printer
        /// </summary>
        /// <param name="filePath">Full path to the file to print</param>
        /// <param name="originalFileName">Original file name for error messages</param>
        /// <returns>PrintResult with status and error information</returns>
        private PrintResult PrintFile(string filePath, string originalFileName)
        {
            try
            {
                // Check if file exists
                if (!File.Exists(filePath))
                {
                    return new PrintResult
                    {
                        Success = false,
                        ErrorMessage = "File not found"
                    };
                }

                // Get file extension to determine handling
                string extension = Path.GetExtension(filePath).ToLower();
                
                // List of extensions that typically don't support printing
                string[] unsupportedExtensions = { ".zip", ".rar", ".7z", ".exe", ".dll", ".msi", 
                    ".mp4", ".avi", ".mov", ".mp3", ".wav", ".flv", ".wmv" };
                
                if (Array.IndexOf(unsupportedExtensions, extension) >= 0)
                {
                    return new PrintResult
                    {
                        Success = false,
                        IsUnsupportedType = true,
                        ErrorMessage = "Unsupported file type"
                    };
                }

                // Use the ShellExecute API to print the file
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    CreateNoWindow = true,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                    UseShellExecute = true
                };

                System.Diagnostics.Process process = System.Diagnostics.Process.Start(psi);
                
                // Wait for the print job to be queued
                // Larger files need more time
                if (process != null)
                {
                    // Calculate wait time based on file size
                    FileInfo fileInfo = new FileInfo(filePath);
                    long fileSizeKB = fileInfo.Length / 1024;
                    
                    // Base wait time of 2 seconds, plus 1 second per MB (max 10 seconds)
                    int waitTime = Math.Min(10000, 2000 + (int)(fileSizeKB));
                    
                    // Wait for process to complete or timeout
                    bool exited = process.WaitForExit(waitTime);
                    
                    // Try to close the process gracefully if it's still running
                    if (!process.HasExited)
                    {
                        try
                        {
                            process.CloseMainWindow();
                            process.WaitForExit(2000); // Give it 2 more seconds
                            
                            if (!process.HasExited)
                            {
                                process.Kill();
                            }
                        }
                        catch
                        {
                            // Process may have already exited
                        }
                    }
                    
                    process.Dispose();
                }

                return new PrintResult { Success = true };
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                // This often means no application is associated with the file type
                return new PrintResult
                {
                    Success = false,
                    IsUnsupportedType = true,
                    ErrorMessage = "No application associated with this file type"
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error printing file: {ex.Message}");
                return new PrintResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// Sanitizes a filename to ensure it's valid for the file system
        /// </summary>
        /// <param name="fileName">Original filename</param>
        /// <returns>Sanitized filename</returns>
        private string SanitizeFileName(string fileName)
        {
            // Remove invalid characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = fileName;
            
            foreach (char c in invalidChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }
            
            // Ensure filename isn't too long (Windows has a 260 char path limit)
            if (sanitized.Length > 200)
            {
                string extension = Path.GetExtension(sanitized);
                sanitized = sanitized.Substring(0, 200 - extension.Length) + extension;
            }
            
            return sanitized;
        }

        /// <summary>
        /// Cleans up temporary directory with retry logic
        /// </summary>
        /// <param name="tempDir">Directory to clean up</param>
        private void CleanupTempDirectory(string tempDir)
        {
            if (!Directory.Exists(tempDir))
                return;

            // Try to delete the directory multiple times
            // Sometimes files are still locked by printing processes
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    Directory.Delete(tempDir, true);
                    return; // Success
                }
                catch
                {
                    if (i < 2) // Don't sleep on last attempt
                    {
                        System.Threading.Thread.Sleep(500);
                    }
                }
            }

            // If we still can't delete it, try to delete individual files
            try
            {
                string[] files = Directory.GetFiles(tempDir);
                foreach (string file in files)
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        // Ignore individual file errors
                    }
                }
                
                // Try to delete directory again
                Directory.Delete(tempDir, false);
            }
            catch
            {
                // If cleanup fails, it's not critical - temp folder will be cleaned up eventually
                System.Diagnostics.Debug.WriteLine($"Warning: Could not fully clean up temp directory: {tempDir}");
            }
        }
    }
}
