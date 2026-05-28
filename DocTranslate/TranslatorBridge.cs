using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace DocTranslate;

public class TranslatorBridge : IDisposable
{
    private Process?      _process;
    private StreamWriter? _stdin;
    private readonly ConcurrentDictionary<string, TaskCompletionSource<string>>                   _pending    = new();
    private readonly ConcurrentDictionary<string, TaskCompletionSource<JsonArray>>               _cmdPending = new();
    private readonly ConcurrentDictionary<string, TaskCompletionSource<(string Text, string Translated)>> _ocrPending = new();
    private int  _idCounter;
    private int  _ocrCounter;
    private bool _ready;
    public  bool IsReady => _ready;

    public async Task<bool> StartAsync(string pythonExe, string scriptPath)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName               = pythonExe,
                Arguments              = $"\"{scriptPath}\"",
                UseShellExecute        = false,
                RedirectStandardInput  = true,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardInputEncoding  = Encoding.UTF8,
            };
            _process = Process.Start(psi) ?? throw new Exception("Failed to start Python.");
            _stdin   = _process.StandardInput;
            _ = Task.Run(ReadLoopAsync);
            _ = Task.Run(() => _process.StandardError.ReadToEndAsync());
            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            while (!_ready && !cts.Token.IsCancellationRequested)
                await Task.Delay(50, cts.Token).ContinueWith(_ => { });
            return _ready;
        }
        catch { return false; }
    }

    private async Task ReadLoopAsync()
    {
        if (_process?.StandardOutput is not { } stdout) return;
        try
        {
            while (await stdout.ReadLineAsync() is { } line)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try
                {
                    var node = JsonNode.Parse(line);
                    if (node is not JsonObject obj) continue;
                    if (obj["ready"]?.GetValue<bool>() == true) { _ready = true; continue; }
                    if (obj["id"] is { } idNode)
                    {
                        var id = idNode.GetValue<string>();
                        // OCR responses carry "text"+"translated"; text-translate responses carry "result"
                        if (id.StartsWith("ocr_"))
                        {
                            if (_ocrPending.TryRemove(id, out var ocrtcs))
                            {
                                var text  = obj["text"]?.GetValue<string>()       ?? "";
                                var trans = obj["translated"]?.GetValue<string>() ?? "";
                                ocrtcs.TrySetResult((text, trans));
                            }
                            continue;
                        }
                        if (_pending.TryRemove(id, out var tcs))
                        {
                            if (obj["error"] is { } err)
                                tcs.TrySetException(new Exception(err.GetValue<string>()));
                            else
                                tcs.TrySetResult(obj["result"]?.GetValue<string>() ?? "");
                        }
                        continue;
                    }
                    if (obj["pairs"]  is JsonArray pairs) { if (_cmdPending.TryRemove("installed_pairs",    out var t1)) t1.TrySetResult(pairs);  continue; }
                    if (obj["packages"] is JsonArray pkgs){ if (_cmdPending.TryRemove("available_packages", out var t2)) t2.TrySetResult(pkgs);   continue; }
                    if (obj["ok"]     is { } ok)          { if (_cmdPending.TryRemove("install",            out var t3)) t3.TrySetResult(new JsonArray(ok.GetValue<bool>())) ; continue; }
                    if (obj["pong"]   is not null)         { if (_cmdPending.TryRemove("ping",              out var t4)) t4.TrySetResult(new JsonArray()); continue; }
                }
                catch { }
            }
        }
        catch { }
    }

    public async Task<string> TranslateAsync(string text, string from, string to,
                                              CancellationToken ct = default)
    {
        if (!_ready || string.IsNullOrWhiteSpace(text)) return text;
        var id  = Interlocked.Increment(ref _idCounter).ToString();
        var tcs = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        _pending[id] = tcs;
        await _stdin!.WriteLineAsync(JsonSerializer.Serialize(new { id, text, from, to }));
        await _stdin.FlushAsync();
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(TimeSpan.FromSeconds(60));
        cts.Token.Register(() => { _pending.TryRemove(id, out _); tcs.TrySetResult(text); });
        return await tcs.Task;
    }

    private async Task<JsonArray> SendCommandAsync(string cmd, object? extra = null, int timeoutMs = 30000)
    {
        if (!_ready) return new JsonArray();
        var tcs = new TaskCompletionSource<JsonArray>(TaskCreationOptions.RunContinuationsAsynchronously);
        _cmdPending[cmd] = tcs;
        string msg;
        if (extra is null)
            msg = JsonSerializer.Serialize(new { cmd });
        else
        {
            var dict = new Dictionary<string, object> { ["cmd"] = cmd };
            foreach (var prop in extra.GetType().GetProperties())
                dict[prop.Name.ToLower()] = prop.GetValue(extra)!;
            msg = JsonSerializer.Serialize(dict);
        }
        await _stdin!.WriteLineAsync(msg);
        await _stdin.FlushAsync();
        var cts = new CancellationTokenSource(timeoutMs);
        cts.Token.Register(() => { _cmdPending.TryRemove(cmd, out _); tcs.TrySetResult(new JsonArray()); });
        return await tcs.Task;
    }

    public async Task<List<(string From, string To)>> GetInstalledPairsAsync()
    {
        var arr    = await SendCommandAsync("installed_pairs");
        var result = new List<(string, string)>();
        foreach (var item in arr)
            if (item is JsonArray pair && pair.Count == 2)
                result.Add((pair[0]!.GetValue<string>(), pair[1]!.GetValue<string>()));
        return result;
    }

    public async Task<List<LangPackage>> GetAvailablePackagesAsync()
    {
        var arr    = await SendCommandAsync("available_packages", timeoutMs: 60000);
        var result = new List<LangPackage>();
        foreach (var item in arr)
            if (item is JsonObject obj)
                result.Add(new LangPackage(
                    obj["from"]?.GetValue<string>() ?? "",
                    obj["to"]?.GetValue<string>()   ?? "",
                    obj["from_name"]?.GetValue<string>() ?? "",
                    obj["to_name"]?.GetValue<string>()   ?? ""));
        return result;
    }

    public async Task<bool> InstallPairAsync(string from, string to)
    {
        var arr = await SendCommandAsync("install", new { from, to }, timeoutMs: 300000);
        return arr.Count > 0 && arr[0]?.GetValue<bool>() == true;
    }

    /// <summary>
    /// OCR an image (supplied as raw bytes) and translate the recognised text.
    /// Returns (original OCR text, translated text). Empty strings on failure.
    /// </summary>
    public async Task<(string Text, string Translated)> OcrImageAsync(
        byte[] imageBytes, string[] fromLangs, string to,
        CancellationToken ct = default)
    {
        if (!_ready || imageBytes.Length == 0) return ("", "");
        var id  = $"ocr_{Interlocked.Increment(ref _ocrCounter)}";
        var tcs = new TaskCompletionSource<(string, string)>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        _ocrPending[id] = tcs;
        var b64 = Convert.ToBase64String(imageBytes);
        var msg = JsonSerializer.Serialize(new {
            cmd = "ocr", id, image_b64 = b64, from_langs = fromLangs, to });
        await _stdin!.WriteLineAsync(msg);
        await _stdin.FlushAsync();
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(TimeSpan.FromSeconds(120));
        cts.Token.Register(() => { _ocrPending.TryRemove(id, out _); tcs.TrySetResult(("", "")); });
        return await tcs.Task;
    }

    /// <summary>
    /// Rasterise a PDF page via Python/PyMuPDF, OCR it, then translate.
    /// Returns (original OCR text, translated text). Empty strings on failure.
    /// </summary>
    public async Task<(string Text, string Translated)> OcrPdfPageAsync(
        string pdfPath, int pageIndex, string[] fromLangs, string to,
        CancellationToken ct = default)
    {
        if (!_ready) return ("", "");
        var id  = $"ocr_{Interlocked.Increment(ref _ocrCounter)}";
        var tcs = new TaskCompletionSource<(string, string)>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        _ocrPending[id] = tcs;
        var msg = JsonSerializer.Serialize(new {
            cmd = "ocr_pdf_page", id, path = pdfPath, page = pageIndex,
            from_langs = fromLangs, to });
        await _stdin!.WriteLineAsync(msg);
        await _stdin.FlushAsync();
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(TimeSpan.FromSeconds(120));
        cts.Token.Register(() => { _ocrPending.TryRemove(id, out _); tcs.TrySetResult(("", "")); });
        return await tcs.Task;
    }

    public void Dispose()
    {
        try { _stdin?.Close();    } catch { }
        try { _process?.Kill();   } catch { }
        _process?.Dispose();
    }
}

public record LangPackage(string FromCode, string ToCode, string FromName, string ToName);
