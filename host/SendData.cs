using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Faktura;

/// <summary>"Send data for testing": packs data.json into a zip and uploads it over HTTPS straight from the program — no mail program needed.
/// The receiving end is a small Google Apps Script web app of the developer (tools\receive-data.gs) that e-mails the file to him,
/// so no password or mail account is stored in this program.</summary>
static class SendData
{
    /// <summary>The address of the receiving web app (…/macros/s/…/exec). Empty = not set up yet.</summary>
    public const string Endpoint = "https://script.google.com/macros/s/AKfycbxiPCvkqhNPozKGvemeyjdt3K6OV1a1hdVs2M62SC00jdivTovDo6ZPCX9EH_UQdUtb/exec";

    static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(90) };

    public static string FolderPath => Path.Combine(Path.GetTempPath(), "invoice-testdata");

    /// <summary>Creates the zip. Returns its path.</summary>
    public static string MakeZip(string stateJson, string info)
    {
        var dir = FolderPath; Directory.CreateDirectory(dir);
        foreach (var old in new DirectoryInfo(dir).GetFiles("invoice-data-*.zip")) if ((DateTime.Now - old.LastWriteTime).TotalDays > 2) try { old.Delete(); } catch { }
        var zip = Path.Combine(dir, $"invoice-data-{DateTime.Now:yyyyMMdd-HHmmss}.zip");
        using (var z = ZipFile.Open(zip, ZipArchiveMode.Create))
        {
            var e = z.CreateEntry("data.json", CompressionLevel.Optimal);
            using (var s = new StreamWriter(e.Open(), new UTF8Encoding(false))) s.Write(stateJson);
            var i = z.CreateEntry("info.txt", CompressionLevel.Optimal);
            using (var s = new StreamWriter(i.Open(), new UTF8Encoding(true))) s.Write(info);
        }
        return zip;
    }

    /// <summary>Only the developer's Google web app (or, in tests, a local address) may receive the file.</summary>
    public static bool Allowed(string url, bool test)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var u)) return false;
        if (test && u.IsLoopback) return true;
        return u.Scheme == Uri.UriSchemeHttps && (u.Host == "script.google.com" || u.Host.EndsWith(".googleusercontent.com"));
    }

    /// <summary>Uploads the zip. Throws with a readable message if it did not arrive.</summary>
    public static async Task Upload(string zip, string info, string version, string url)
    {
        var payload = JsonSerializer.Serialize(new { app = "invoice", version, info, name = Path.GetFileName(zip), zip = Convert.ToBase64String(File.ReadAllBytes(zip)) });
        using var content = new StringContent(payload, Encoding.UTF8, "text/plain");
        string body;
        try
        {
            using var resp = await Http.PostAsync(url, content);
            body = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode) throw new Exception("Сървърът отговори с код " + (int)resp.StatusCode + ".");
        }
        catch (HttpRequestException) { throw new Exception("Няма връзка с интернет или сървърът не отговаря."); }
        catch (TaskCanceledException) { throw new Exception("Сървърът не отговори навреме."); }
        JsonNode j;
        try { j = JsonNode.Parse(body); } catch (JsonException) { throw new Exception("Неочакван отговор от сървъра."); }
        if (j?["ok"]?.GetValue<bool>() != true) throw new Exception("Сървърът не прие файла" + (j?["error"] != null ? " (" + j["error"] + ")" : "") + ".");
    }
}
