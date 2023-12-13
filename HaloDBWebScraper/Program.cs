using HtmlAgilityPack;

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Linq;
using OfficeOpenXml;
using static OfficeOpenXml.ExcelErrorValue;

class Program
{
    static void Main()
    {
        // Base URL for the leaderboard
        string baseUrl = "https://halotracker.com/halo-infinite/leaderboards/csr/all/default?page=";

        List<PlayerData> playersData = new List<PlayerData>();

        // Titles to select (customize as needed)
        List<string> selectedStatTitles = new List<string> { "Win %", "Avg KDA", "K/D Ratio","Avg. Damage","Kills","Assists","Deaths","Headshots","Matches Won","Matches Lost"/* Add more titles as needed */ };

        // Iterate through pages (up to page 10)
        for (int page = 1; page <= 10; page++)
        {
            // Construct the URL for the current page
            string url = $"{baseUrl}{page}&playlist=1";

            // Use HtmlWeb to load the HTML content of the webpage
            var web = new HtmlWeb();
            var doc = web.Load(url);

            // Extract player names and links to their stats pages
            var playerNodes = doc.DocumentNode.SelectNodes("//span[@class='trn-ign__username']");

            if (playerNodes != null)
            {
                foreach (var playerNode in playerNodes)
                {
                    string playerName = playerNode.InnerText.Trim();

                    // Construct the stats page link for the current player
                    string statsPageLink = ConstructStatsPageLink(playerName);

                    // Visit the stats page and extract player stats
                    var statsPage = web.Load(statsPageLink);
                    var playerStats = ExtractPlayerStats(statsPage.DocumentNode, selectedStatTitles);

                    // Process or store player data as needed
                    Console.WriteLine($"Player: {playerName}");
                    Console.WriteLine($"Stats Page Link: {statsPageLink}");
                    foreach (var stat in playerStats)
                    {
                        Console.WriteLine($"  Stat Title: {stat.Title}, Stat Value: {stat.Value}");
                    }

                    // Add player data to the list
                    playersData.Add(new PlayerData { PlayerName = playerName, Stats = playerStats });
                }
            }

            // Export player data to Excel inside the loop
            ExportToExcel(playersData, selectedStatTitles);
        }
    }

    static string ConstructStatsPageLink(string playerName)
    {
        // Encode the player name for URL
        string encodedPlayerName = Uri.EscapeDataString(playerName);

        // Construct the stats page link
        return $"https://halotracker.com/halo-infinite/profile/xbl/{encodedPlayerName}/overview?experience=ranked&playlist=edfef3ac-9cbe-4fa2-b949-8f29deafd483";
    }


    static List<StatData> ExtractPlayerStats(HtmlNode playerNode, List<string> selectedStatTitles)
    {
        List<StatData> playerStats = new List<StatData>();
        try
        {
            var mmrContainerNode = playerNode.SelectNodes(".//ancestor::div[@class='stat']//span[@class='stat__value']");

            if (mmrContainerNode != null && mmrContainerNode.Count > 0)
            {
                string mmrRating = mmrContainerNode[0]?.InnerText.Trim();
                playerStats.Add(new StatData { Title = "MMR Rating", Value = mmrRating });

                // Print the MMR stat and its value for testing
                Console.WriteLine($"MMR Title: MMR Rating, MMR Value: {mmrRating}");
            }
            else
            {
                Console.WriteLine("MMR Rating not found in the expected structure.");
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"Error extracting MMR rating: {e.Message}");
        }
        try
        {
            // Assuming 'playerNode' is the span containing the player name
            var statsContainerNodes = playerNode.SelectNodes(".//ancestor::div[@class='giant-stats']//div[@class='stat align-left giant expandable']");

            if (statsContainerNodes != null)
            {
                foreach (var statsContainerNode in statsContainerNodes)
                {
                    string statTitle = statsContainerNode.SelectSingleNode(".//span[@class='name']")?.InnerText.Trim();
                    string statValue = statsContainerNode.SelectSingleNode(".//span[@class='value']")?.InnerText.Trim();

                    // Check if the stat title is in the selected list
                    if (selectedStatTitles.Contains(statTitle))
                    {
                        // Add the stat to the list
                        playerStats.Add(new StatData { Title = statTitle, Value = statValue });

                        // Print the stat and its value for testing
                        Console.WriteLine($"Stat Title: {statTitle}, Stat Value: {statValue}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log or print the exception details for debugging
            Console.WriteLine($"Error extracting stats: {ex.Message}");
        }
        try
        {
            //
            // Assuming 'playerNode' is the span containing the player name
            var statsContainerNodes = playerNode.SelectNodes(".//ancestor::div[@class='main']//div[@class='stat align-left expandable']");

            if (statsContainerNodes != null)
            {
                foreach (var statsContainerNode in statsContainerNodes)
                {
                    string statTitle = statsContainerNode.SelectSingleNode(".//span[@class='name']")?.InnerText.Trim();
                    string statValue = statsContainerNode.SelectSingleNode(".//span[@class='value']")?.InnerText.Trim();

                    // Check if the stat title is in the selected list
                    if (selectedStatTitles.Contains(statTitle))
                    {
                        // Add the stat to the list
                        playerStats.Add(new StatData { Title = statTitle, Value = statValue });

                        // Print the stat and its value for testing
                        Console.WriteLine($"Stat Title: {statTitle}, Stat Value: {statValue}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log or print the exception details for debugging
            Console.WriteLine($"Error extracting stats: {ex.Message}");
        }

        return playerStats;
    }


    static void ExportToExcel(List<PlayerData> playersData, List<string> selectedStatTitles)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            // Add a worksheet to the workbook
            var worksheet = package.Workbook.Worksheets.Add("Player Stats");

            // Headers
            worksheet.Cells[1, 1].Value = "Player Name";
            int column = 2;
            foreach (var statTitle in selectedStatTitles)
            {
                worksheet.Cells[1, column].Value = statTitle;
                column++;
            }

            // Data
            for (int row = 2; row <= playersData.Count + 1; row++)
            {
                worksheet.Cells[row, 1].Value = playersData[row - 2].PlayerName;

                column = 2;
                foreach (var statTitle in selectedStatTitles)
                {
                    var stat = playersData[row - 2].Stats.Find(s => s.Title == statTitle);
                    worksheet.Cells[row, column].Value = stat?.Value;
                    column++;
                }
            }

            // Save the Excel file
            package.SaveAs(new FileInfo("PlayerStats.xlsx"));
        }

        Console.WriteLine("Player data exported to Excel successfully.");
    }
}

class PlayerData
{
    public string PlayerName { get; set; }
    public List<StatData> Stats { get; set; }
}

class StatData
{
    public string Title { get; set; }
    public string Value { get; set; }
}
