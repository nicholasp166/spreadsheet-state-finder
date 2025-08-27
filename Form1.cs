namespace cswf;

using System;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Text;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.Json;
using System.Net;
using dotenv.net;
using dotenv.net.Utilities;

public partial class Form1 : Form
{

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern bool AllocConsole();
    private WebBrowser fileContentsBrowser;

    // Global request counter
    private int requestCount =4998;
    private const int maxRequests = 4999;

    public Form1()
    {
        AllocConsole();
        InitializeComponent();
        DotEnv.Load();
        
        // Create and add a WebBrowser to display file contents as HTML
        fileContentsBrowser = new WebBrowser();
        fileContentsBrowser.Location = new System.Drawing.Point(50, 250);
        fileContentsBrowser.Size = new System.Drawing.Size(700, 150);
        this.Controls.Add(fileContentsBrowser);

        // Set EPPlus license for EPPlus 8 and later
        ExcelPackage.License.SetNonCommercialPersonal("test");
    }

    private async Task<string> GetRequestAsync(string url)
    {
        if (requestCount >= maxRequests)
        {
            Console.WriteLine("Maximum number of requests reached. Further requests are blocked.");
            throw new InvalidOperationException("Maximum number of requests reached. Further requests are blocked.");
        }

        requestCount++; // Increment the counter each time the function is called

        using (HttpClient client = new HttpClient())
        {
            HttpResponseMessage response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
    }
    
    private async Task PrintGetRequestToConsole(string url)
    {
        try
        {
            string response = await GetRequestAsync(url);

            Console.WriteLine("\nGet Response:");
            //Console.WriteLine(response + "\n");
            //Console.WriteLine($"{response}" + Environment.NewLine);
            using JsonDocument doc = JsonDocument.Parse(response);
            Console.WriteLine(response);
            JsonElement root = doc.RootElement;
            //Console.WriteLine(root.GetProperty("display_name").GetString());
            foreach (JsonElement item in root.EnumerateArray())
            {
                string? dn = item.GetProperty("display_name").GetString() != null ? item.GetProperty("display_name").GetString() : "No display found";
                if (dn != null)
                {
                    string[] dnA = dn.Split(", ");
                    string zipcode = dnA[5];
                }

            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during GET: " + ex.Message);
        }
    }

    private async Task UpdateSpreadsheetWithZipcode(string url, string filePath)
    {
        try
        {
            string response = await GetRequestAsync(url);

            using JsonDocument doc = JsonDocument.Parse(response);
            JsonElement root = doc.RootElement;
            string zipcode = "N/A";
            foreach (JsonElement item in root.EnumerateArray())
            {
                string? dn = item.GetProperty("display_name").GetString();
                if (!string.IsNullOrEmpty(dn))
                {
                    string[] dnA = dn.Split(", ");
                    if (dnA.Length > 5)
                        zipcode = dnA[5];
                    else
                        zipcode = "N/A";
                    Console.WriteLine(zipcode);
                }
            }

            // Block to open the Excel file and write the zipcode to the appropriate row
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rows = worksheet.Dimension.Rows;
                int cols = worksheet.Dimension.Columns;

                // Find or create the "zip" column
                int zipColIndex = -1;
                for (int c = 1; c <= cols; c++)
                {
                    string header = worksheet.Cells[1, c].Text.Trim().ToLower();
                    if (header == "zip")
                    {
                        zipColIndex = c;
                        break;
                    }
                }
                if (zipColIndex == -1)
                {
                    zipColIndex = cols + 1;
                    worksheet.Cells[1, zipColIndex].Value = "zip";
                }

                // Find the next empty row for zipcode (assuming you want to update the last row processed)
                // If you want to update a specific row, pass the row index as a parameter
                for (int r = 2; r <= rows; r++)
                {
                    // Only update if the cell is empty
                    if (string.IsNullOrEmpty(worksheet.Cells[r, zipColIndex].Text))
                    {
                        worksheet.Cells[r, zipColIndex].Value = zipcode;
                        Console.WriteLine($"Wrote zipcode '{zipcode}' to row {r} in column {zipColIndex}");
                        break; // Write to the first empty cell found
                    }
                }

                package.Save();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during GET: " + ex.Message);
        }
    }

    private async void uploadButton_Click(object sender, EventArgs e)
    {
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            string filePath = openFileDialog.FileName;
            MessageBox.Show($"Selected file: {filePath}", "File Uploaded", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Display contents if CSV
            if (Path.GetExtension(filePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                string[] lines = File.ReadAllLines(filePath);
                var sb = new StringBuilder();
                sb.Append("<html><body><table border='1' style='border-collapse:collapse;'>");

                // Parse CSV into array for console output
                var csvData = lines.Select(line => line.Split(',')).ToArray();

                // Find max width for each column
                int colCount = csvData.Max(arr => arr.Length);
                int[] colWidths = new int[colCount];
                foreach (var row in csvData)
                {
                    for (int i = 0; i < row.Length; i++)
                    {
                        int len = row[i].Length;
                        if (len > colWidths[i])
                            colWidths[i] = len;
                    }
                }

                foreach (var line in lines)
                {
                    sb.Append("<tr>");
                    foreach (var cell in line.Split(','))
                    {
                        sb.AppendFormat("<td>{0}</td>", System.Net.WebUtility.HtmlEncode(cell));
                    }
                    sb.Append("</tr>");
                }
                sb.Append("</table></body></html>");
                fileContentsBrowser.DocumentText = sb.ToString();
            }
            // Display contents if XLSX
            else if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        int rows = worksheet.Dimension.Rows;
                        int cols = worksheet.Dimension.Columns;
                        var sb = new StringBuilder();
                        Dictionary<int, string> addrs = new Dictionary<int, string>();
                        sb.Append("<html><body><table border='1' style='border-collapse:collapse;'>");

                        //didnt notice this starts at one, this is actually cursed
                        for (int r = 1; r <= rows; r++)
                        {
                            sb.Append("<tr>");
                            for (int c = 1; c <= cols; c++)
                            {
                                string cellText = worksheet.Cells[r, c].Text;
                                sb.AppendFormat("<td>{0}</td>", System.Net.WebUtility.HtmlEncode(cellText));


                                if (r > 1)
                                {
                                    //add dictionary here to parse address info
                                    //Console.WriteLine(r);
                                    //Console.WriteLine(System.Net.WebUtility.HtmlEncode(cellText) + " ");
                                    if (addrs.ContainsKey(r))
                                    {
                                        addrs[r] = addrs[r] + " " + System.Net.WebUtility.HtmlEncode(cellText) + " ";
                                    }
                                    else
                                    {
                                        addrs[r] = System.Net.WebUtility.HtmlEncode(cellText) + " ";
                                    }
                                }

                            }
                            sb.Append("</tr>");
                        }
                        string url = "";
                        string fullLink = EnvReader.GetStringValue("LOCATIONIQ_API_LINK");
                        string apiKey = EnvReader.GetStringValue("LOCATIONIQ_API_KEY");
                        //main loop going throug heach site
                        foreach (KeyValuePair<int, string> entry in addrs)
                        {
                            url = fullLink + apiKey + "&q=" + WebUtility.UrlEncode(entry.Value) + "&format=json";
                            //Console.WriteLine(entry.Value);
                            //Console.WriteLine(url);
                            //await PrintGetRequestToConsole(url);
                            await UpdateSpreadsheetWithZipcode(url, filePath);
                            await Task.Delay(2000);
                        }
                        sb.Append("</table></body></html>");
                        fileContentsBrowser.DocumentText = sb.ToString();
                    }
                }
                catch (Exception ex)
                {
                    fileContentsBrowser.DocumentText = $"<html><body>Error reading Excel file: {System.Net.WebUtility.HtmlEncode(ex.Message)}</body></html>";
                }
            }
            else
            {
                fileContentsBrowser.DocumentText = "<html><body>Preview for this file type is not supported.</body></html>";
            }
        }
    }
}
