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
using System.Net;
using dotenv.net;
using dotenv.net.Utilities;

public partial class Form1 : Form
{

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern bool AllocConsole();
    private WebBrowser fileContentsBrowser;
    

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
            Console.WriteLine("Get Response:");
            Console.WriteLine();
            Console.WriteLine(response);
            Console.WriteLine();
            Console.WriteLine();
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
                        foreach (KeyValuePair<int, string> entry in addrs)
                        {
                            url = fullLink+ apiKey + "&q=" + WebUtility.UrlEncode(entry.Value) + "&format=json";
                            Console.WriteLine(entry.Value);
                            Console.WriteLine(url);
                            await PrintGetRequestToConsole(url);
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
