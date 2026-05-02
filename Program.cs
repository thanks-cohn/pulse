// Pulse — tiny life logger
//
// Install:
//   dotnet publish -c Release -r linux-x64 --self-contained true /p:PublishSingleFile=true
//   cp bin/Release/net*/linux-x64/publish/Pulse ~/.local/bin/pulse
//   chmod +x ~/.local/bin/pulse
//
// Storage:
//   ~/.local/share/pulse/events.jsonl
//   ~/.local/share/pulse/state.json
//   ~/.config/pulse/config.json

using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PulseCli;

public sealed class PulseEvent
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("ts")]
    public string Timestamp { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("category")]
    public string Category { get; set; } = "";

    [JsonPropertyName("note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }

    [JsonPropertyName("source")]
    public string Source { get; set; } = "cli";

    [JsonPropertyName("meta")]
    public Dictionary<string, JsonElement> Meta { get; set; } = new();
}

public sealed class PulseState
{
    [JsonPropertyName("open_toggles")]
    public Dictionary<string, OpenToggle> OpenToggles { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class OpenToggle
{
    [JsonPropertyName("event_id")]
    public string EventId { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("category")]
    public string Category { get; set; } = "";

    [JsonPropertyName("started_at")]
    public string StartedAt { get; set; } = "";

    [JsonPropertyName("note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; set; }
}

public sealed class PulseConfig
{
    [JsonPropertyName("sleep_action")]
    public string SleepAction { get; set; } = "none";

    [JsonPropertyName("auto_wake")]
    public bool AutoWake { get; set; } = false;
}

public sealed class ParsedCommand
{
    public string Command { get; init; } = "";
    public DateTimeOffset Timestamp { get; init; }
    public bool HasExplicitTime { get; init; }
    public string? Note { get; init; }
    public List<string> Tags { get; init; } = new();
}

public static class Program
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private static readonly JsonSerializerOptions PrettyJsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private static readonly HashSet<string> ToggleCommands = new(StringComparer.OrdinalIgnoreCase)
    {
        "meal", "breakfast", "lunch", "dinner", "snack", "break", "walk", "work", "exercise", "study"
    };

    private static readonly HashSet<string> PointCommands = new(StringComparer.OrdinalIgnoreCase)
    {
        "wake", "pain", "mood", "energy", "note", "dream", "journal", "important"
    };

    private static string Home => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    private static string ConfigDir => Path.Combine(Home, ".config", "pulse");
    private static string DataDir => Path.Combine(Home, ".local", "share", "pulse");
    private static string ConfigPath => Path.Combine(ConfigDir, "config.json");
    private static string EventsPath => Path.Combine(DataDir, "events.jsonl");
    private static string StatePath => Path.Combine(DataDir, "state.json");

    public static int Main(string[] args)
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            Directory.CreateDirectory(DataDir);
            EnsureDefaultConfig();

            if (args.Length == 0 || IsHelp(args[0]))
            {
                PrintHelp();
                return 0;
            }

            string cmd = NormalizeCommand(args[0]);
            string[] rest = args.Skip(1).ToArray();

            return cmd switch
            {
                "status" => CmdStatus(rest),
                "where" => CmdWhere(rest),
                "config" => CmdConfig(rest),
                "day" => CmdDay(rest),
                "year" => CmdYear(rest),
                "today" => CmdDay(Array.Empty<string>()),
                "yesterday" => CmdDay(new[] { DateOnly.FromDateTime(DateTime.Now.AddDays(-1)).ToString("yyyy-MM-dd") }),
                "tomorrow" => CmdTomorrow(),
                "promise" => CmdPromise(rest),
                "log" => CmdLog(rest),
                "last" => CmdLast(),
                "edit" => CmdEdit(rest),
                "append" => CmdAppend(rest),
                "sleep" => CmdSleep(rest),
                _ when ToggleCommands.Contains(cmd) => CmdToggle(cmd, rest),
                _ when PointCommands.Contains(cmd) => CmdPoint(cmd, rest),
                _ => CmdCustom(cmd, rest)
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"pulse error: {ex.Message}");
            return 1;
        }
    }

    private static bool IsHelp(string arg) => arg is "help" or "--help" or "-h";

    private static bool WantsJson(string[] args)
    {
        return args.Any(a => a.Equals("--json", StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizeCommand(string raw)
    {
        string s = raw.Trim().ToLowerInvariant();
        return s switch
        {
            "slept" => "sleep",
            "woke" => "wake",
            "hurt" => "pain",
            "ache" => "pain",
            "feeling" => "mood",
            "feel" => "mood",
            "dreams" => "dream",
            _ => s
        };
    }

    private static void PrintHelp()
    {
        Console.WriteLine("Pulse — tiny life logger");
        Console.WriteLine();
        Console.WriteLine("Shape:");
        Console.WriteLine("  pulse <event> [time] [\"note\"]");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  pulse wake \"dream journal: old house\"");
        Console.WriteLine("  pulse wake 9:30am \"dream journal: ocean, black dog\"");
        Console.WriteLine("  pulse lunch \"tacos and coffee\"");
        Console.WriteLine("  pulse lunch");
        Console.WriteLine("  pulse pain \"left knee hurts going upstairs\"");
        Console.WriteLine("  pulse mood \"sad but steady\" tag:glowbox");
        Console.WriteLine("  pulse important \"launched comic2panel — first clean push\" tag:github tag:tool");
        Console.WriteLine("  pulse sleep");
        Console.WriteLine("  pulse sleep 2:40am \"couldn't sleep because my mind kept running\"");
        Console.WriteLine();
        Console.WriteLine("Reports:");
        Console.WriteLine("  pulse day");
        Console.WriteLine("  pulse day 2026-05-02");
        Console.WriteLine("  pulse today");
        Console.WriteLine("  pulse yesterday");
        Console.WriteLine("  pulse tomorrow");
        Console.WriteLine("  pulse promise me");
        Console.WriteLine("  pulse year");
        Console.WriteLine("  pulse year 2026");
        Console.WriteLine("  pulse log");
        Console.WriteLine("  pulse log 100");
        Console.WriteLine("  pulse last");
        Console.WriteLine();
        Console.WriteLine("Config:");
        Console.WriteLine("  pulse config");
        Console.WriteLine("  pulse config sleep-action none");
        Console.WriteLine("  pulse config sleep-action shutdown");
        Console.WriteLine("  pulse where");
        Console.WriteLine("  pulse status");
    }

    private static int CmdTomorrow()
    {
        Console.WriteLine("the future is up to us Quandranea <3");
        return 0;
    }

    private static int CmdWhere(string[] rest)
    {
        if (WantsJson(rest))
        {
            var obj = new
            {
                config = ConfigPath,
                events = EventsPath,
                state = StatePath
            };

            Console.WriteLine(JsonSerializer.Serialize(obj, PrettyJsonOptions));
            return 0;
        }

        Console.WriteLine($"config: {ConfigPath}");
        Console.WriteLine($"events: {EventsPath}");
        Console.WriteLine($"state:  {StatePath}");
        return 0;
    }

    private static int CmdStatus(string[] rest)
    {
        var config = LoadConfig();
        var state = LoadState();

        var open = state.OpenToggles.Select(kv => new
        {
            name = kv.Key,
            started_at = kv.Value.StartedAt,
            friendly_started = FriendlyTime(ParseIso(kv.Value.StartedAt)),
            note = kv.Value.Note
        }).ToList();

        if (WantsJson(rest))
        {
            var obj = new
            {
                ok = true,
                events = EventsPath,
                config = ConfigPath,
                state = StatePath,
                sleep_action = config.SleepAction,
                open_toggles = open
            };

            Console.WriteLine(JsonSerializer.Serialize(obj, PrettyJsonOptions));
            return 0;
        }

        Console.WriteLine("Pulse status");
        Console.WriteLine($"events: {EventsPath}");
        Console.WriteLine($"sleep_action: {config.SleepAction}");
        Console.WriteLine();

        if (open.Count == 0)
        {
            Console.WriteLine("open toggles: none");
            return 0;
        }

        Console.WriteLine("open toggles:");
        foreach (var item in open)
            Console.WriteLine($"  {item.name}: started {item.friendly_started}" + NoteSuffix(item.note));

        return 0;
    }

    private static int CmdConfig(string[] rest)
    {
        var config = LoadConfig();

        if (rest.Length == 0 || WantsJson(rest))
        {
            Console.WriteLine(JsonSerializer.Serialize(config, PrettyJsonOptions));
            return 0;
        }

        if (rest.Length >= 2 && rest[0] == "sleep-action")
        {
            string value = rest[1].ToLowerInvariant();

            if (value is not ("none" or "shutdown" or "suspend" or "hibernate"))
            {
                Console.Error.WriteLine("sleep-action must be one of: none, shutdown, suspend, hibernate");
                return 2;
            }

            config.SleepAction = value;
            SaveConfig(config);
            Console.WriteLine($"sleep_action = {value}");
            return 0;
        }

        Console.Error.WriteLine("usage: pulse config sleep-action none|shutdown|suspend|hibernate");
        return 2;
    }

    private static int CmdPoint(string cmd, string[] rest)
    {
        var parsed = ParseCommand(cmd, rest);
        var ev = NewEvent(cmd, CategoryForPoint(cmd), parsed.Timestamp, parsed.Note);
        ApplyTags(ev, parsed.Tags);
        AppendEvent(ev);
        Console.WriteLine($"{Title(cmd)} logged: {FriendlyTime(parsed.Timestamp)}" + NoteSuffix(parsed.Note));
        return 0;
    }

    private static int CmdCustom(string cmd, string[] rest)
    {
        var parsed = ParseCommand(cmd, rest);
        var ev = NewEvent(cmd, "custom", parsed.Timestamp, parsed.Note);
        ApplyTags(ev, parsed.Tags);
        AppendEvent(ev);
        Console.WriteLine($"{Title(cmd)} logged: {FriendlyTime(parsed.Timestamp)}" + NoteSuffix(parsed.Note));
        return 0;
    }

    private static int CmdToggle(string cmd, string[] rest)
    {
        var parsed = ParseCommand(cmd, rest);
        string category = CategoryForToggle(cmd);
        var state = LoadState();

        if (state.OpenToggles.TryGetValue(cmd, out var open))
        {
            var start = ParseIso(open.StartedAt);
            var duration = Math.Max(0, (int)Math.Round((parsed.Timestamp - start).TotalMinutes));

            var ev = NewEvent($"{cmd}_end", category, parsed.Timestamp, parsed.Note);
            ApplyTags(ev, parsed.Tags);
            SetMeta(ev, "duration_minutes", duration);
            SetMeta(ev, "started_at", open.StartedAt);
            SetMeta(ev, "start_event_id", open.EventId);

            AppendEvent(ev);
            state.OpenToggles.Remove(cmd);
            SaveState(state);

            Console.WriteLine($"{Title(cmd)} ended: {FriendlyTime(parsed.Timestamp)} ({FormatMinutes(duration)})" + NoteSuffix(parsed.Note));
            return 0;
        }

        var startEv = NewEvent($"{cmd}_start", category, parsed.Timestamp, parsed.Note);
        ApplyTags(startEv, parsed.Tags);
        AppendEvent(startEv);

        state.OpenToggles[cmd] = new OpenToggle
        {
            EventId = startEv.Id,
            Type = startEv.Type,
            Category = category,
            StartedAt = startEv.Timestamp,
            Note = parsed.Note
        };

        SaveState(state);
        Console.WriteLine($"{Title(cmd)} started: {FriendlyTime(parsed.Timestamp)}" + NoteSuffix(parsed.Note));
        return 0;
    }

    private static int CmdSleep(string[] rest)
    {
        var parsed = ParseCommand("sleep", rest);
        var config = LoadConfig();

        if (parsed.HasExplicitTime)
        {
            var ev = NewEvent("sleep_actual", "sleep", parsed.Timestamp, parsed.Note);
            ApplyTags(ev, parsed.Tags);
            AppendEvent(ev);
            Console.WriteLine($"Actual sleep logged: {FriendlyTime(parsed.Timestamp)}" + NoteSuffix(parsed.Note));
            return 0;
        }

        var attempt = NewEvent("sleep_attempt", "sleep", parsed.Timestamp, parsed.Note);
        ApplyTags(attempt, parsed.Tags);
        SetMeta(attempt, "action", config.SleepAction);
        AppendEvent(attempt, fsync: true);

        Console.WriteLine($"Sleep attempt logged: {FriendlyTime(parsed.Timestamp)}" + NoteSuffix(parsed.Note));

        if (config.SleepAction == "none")
        {
            Console.WriteLine("sleep_action is none. To make 'pulse sleep' shut down: pulse config sleep-action shutdown");
            return 0;
        }

        return RunPowerAction(config.SleepAction);
    }

    private static int RunPowerAction(string action)
    {
        string command = action switch
        {
            "shutdown" => "poweroff",
            "suspend" => "suspend",
            "hibernate" => "hibernate",
            _ => ""
        };

        if (string.IsNullOrWhiteSpace(command))
            return 0;

        Console.WriteLine($"Running: systemctl {command}");

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "systemctl",
                UseShellExecute = false
            };

            psi.ArgumentList.Add(command);

            using var process = Process.Start(psi);
            process?.WaitForExit();
            return process?.ExitCode ?? 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Could not run systemctl {command}: {ex.Message}");
            return 1;
        }
    }

    private static int CmdLog(string[] rest)
    {
        int limit = 30;

        if (rest.Length > 0 && int.TryParse(rest[0], out var parsed))
            limit = Math.Max(1, parsed);

        var events = ReadEvents().TakeLast(limit).ToList();

        if (events.Count == 0)
        {
            Console.WriteLine("No events yet.");
            return 0;
        }

        foreach (var ev in events)
            Console.WriteLine(FormatEventLine(ev));

        return 0;
    }

    private static int CmdLast()
    {
        var ev = ReadEvents().LastOrDefault();

        if (ev is null)
        {
            Console.WriteLine("No events yet.");
            return 0;
        }

        Console.WriteLine(JsonSerializer.Serialize(ev, PrettyJsonOptions));
        return 0;
    }

    private static int CmdEdit(string[] rest)
    {
        if (rest.Length < 2 || rest[0] != "last")
        {
            Console.Error.WriteLine("usage: pulse edit last \"new note\"");
            return 2;
        }

        string note = string.Join(' ', rest.Skip(1)).Trim();

        if (string.IsNullOrWhiteSpace(note))
        {
            Console.Error.WriteLine("missing note");
            return 2;
        }

        return RewriteLastEvent(ev => ev.Note = note, "edited last event");
    }

    private static int CmdAppend(string[] rest)
    {
        if (rest.Length < 2 || rest[0] != "last")
        {
            Console.Error.WriteLine("usage: pulse append last \"extra note\"");
            return 2;
        }

        string note = string.Join(' ', rest.Skip(1)).Trim();

        if (string.IsNullOrWhiteSpace(note))
        {
            Console.Error.WriteLine("missing note");
            return 2;
        }

        return RewriteLastEvent(ev =>
        {
            ev.Note = string.IsNullOrWhiteSpace(ev.Note) ? note : ev.Note + " " + note;
        }, "appended to last event");
    }

    private static int RewriteLastEvent(Action<PulseEvent> edit, string message)
    {
        var events = ReadEvents().ToList();

        if (events.Count == 0)
        {
            Console.WriteLine("No events yet.");
            return 0;
        }

        edit(events[^1]);
        WriteAllEvents(events);
        Console.WriteLine(message);
        return 0;
    }

    private static int CmdDay(string[] rest)
    {
        DateOnly day = DateOnly.FromDateTime(DateTime.Now);

        if (rest.Length > 0)
        {
            string raw = rest[0].Trim();

            if (!TryParseDay(raw, out day))
            {
                Console.Error.WriteLine("usage: pulse day [yyyy-MM-dd]");
                Console.Error.WriteLine("example: pulse day 2026-05-02");
                return 2;
            }
        }

        var events = ReadEvents()
            .Where(e => TryParseIso(e.Timestamp, out var ts) && DateOnly.FromDateTime(ts.LocalDateTime) == day)
            .OrderBy(e => ParseIso(e.Timestamp))
            .ToList();

        Console.WriteLine($"Pulse — {day:MMMM d, yyyy}");
        Console.WriteLine();

        if (events.Count == 0)
        {
            Console.WriteLine("No events for this day.");
            return 0;
        }

        PrintPrettyDayEvents(events);

        Console.WriteLine();
        PrintCompactDaySummary(events);
        return 0;
    }

    private static void PrintPrettyDayEvents(List<PulseEvent> events)
    {
        DateTimeOffset? lastTime = null;
        int noteWaveIndex = 0;
        int[] noteWave = { 0, 1, 2, 3, 4, 5, 4, 3, 2, 1 };

        foreach (var ev in events)
        {
            var ts = ParseIso(ev.Timestamp);

            if (lastTime is not null)
            {
                double hours = (ts - lastTime.Value).TotalHours;
                int blankLines = Math.Min(4, Math.Max(0, (int)Math.Floor(hours)));

                for (int i = 0; i < blankLines; i++)
                    Console.WriteLine();
            }

            int indent = 0;

            if (ev.Type.Equals("note", StringComparison.OrdinalIgnoreCase))
            {
                indent = noteWave[noteWaveIndex % noteWave.Length];
                noteWaveIndex++;
            }
            else
            {
                noteWaveIndex = 0;
            }

            PrintEventWrapped(ev, indent);
            lastTime = ts;
        }
    }

    private static void PrintEventWrapped(PulseEvent ev, int indent)
    {
        var ts = ParseIso(ev.Timestamp);
        string time = FriendlyTime(ts).PadLeft(8);
        string label = HumanType(ev.Type).PadRight(10);
        string duration = TryGetMetaInt(ev, "duration_minutes", out var mins)
            ? $" ({FormatMinutes(mins)})"
            : "";
        string tags = GetMetaStringArrayDisplay(ev, "tags") is { Length: > 0 } tagText
            ? $"  [{tagText}]"
            : "";

        string note = string.IsNullOrWhiteSpace(ev.Note) ? tags.TrimStart() : ev.Note + tags;
        string leftIndent = new(' ', indent);
        string prefix = $"{time}  {label}{duration} ";
        string continuationPrefix = leftIndent + new string(' ', prefix.Length);

        int terminalWidth;
        try
        {
            terminalWidth = Console.WindowWidth > 0 ? Console.WindowWidth : 80;
        }
        catch
        {
            terminalWidth = 80;
        }

        int noteWidth = Math.Max(20, terminalWidth - leftIndent.Length - prefix.Length - 2);
        var lines = WrapTextPreservingParagraphs(note, noteWidth);

        if (lines.Count == 0)
        {
            Console.WriteLine(leftIndent + prefix.TrimEnd());
            return;
        }

        Console.WriteLine(leftIndent + prefix + lines[0]);

        for (int i = 1; i < lines.Count; i++)
        {
            if (lines[i].Length == 0)
                Console.WriteLine();
            else
                Console.WriteLine(continuationPrefix + lines[i]);
        }
    }

    private static List<string> WrapTextPreservingParagraphs(string text, int maxWidth)
    {
        var lines = new List<string>();

        if (string.IsNullOrWhiteSpace(text))
            return lines;

        string normalized = text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace("\r", "\n", StringComparison.Ordinal);

        string[] paragraphs = normalized.Split("\n\n", StringSplitOptions.None);

        foreach (var rawParagraph in paragraphs)
        {
            string paragraph = rawParagraph.Trim();

            if (paragraph.Length == 0)
            {
                if (lines.Count > 0 && lines[^1].Length != 0)
                    lines.Add("");

                continue;
            }

            var wrapped = WrapSingleParagraph(paragraph, maxWidth);
            lines.AddRange(wrapped);

            // Preserve one quiet blank line between paragraphs.
            lines.Add("");
        }

        while (lines.Count > 0 && lines[^1].Length == 0)
            lines.RemoveAt(lines.Count - 1);

        return lines;
    }

    private static List<string> WrapSingleParagraph(string text, int maxWidth)
    {
        var lines = new List<string>();

        if (string.IsNullOrWhiteSpace(text))
            return lines;

        var words = text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var current = new StringBuilder();

        foreach (var word in words)
        {
            if (word.Length > maxWidth)
            {
                if (current.Length > 0)
                {
                    lines.Add(current.ToString());
                    current.Clear();
                }

                for (int i = 0; i < word.Length; i += maxWidth)
                    lines.Add(word.Substring(i, Math.Min(maxWidth, word.Length - i)));

                continue;
            }

            if (current.Length == 0)
            {
                current.Append(word);
                continue;
            }

            if (current.Length + 1 + word.Length > maxWidth)
            {
                lines.Add(current.ToString());
                current.Clear();
                current.Append(word);
            }
            else
            {
                current.Append(' ');
                current.Append(word);
            }
        }

        if (current.Length > 0)
            lines.Add(current.ToString());

        return lines;
    }

    private static void PrintCompactDaySummary(List<PulseEvent> events)
    {
        int notes = events.Count(e => e.Type.Equals("note", StringComparison.OrdinalIgnoreCase));
        int moods = events.Count(e => e.Type.Equals("mood", StringComparison.OrdinalIgnoreCase));
        int important = events.Count(e => e.Type.Equals("important", StringComparison.OrdinalIgnoreCase)
            || e.Category.Equals("important", StringComparison.OrdinalIgnoreCase));

        Console.WriteLine("Summary");
        Console.WriteLine($"{events.Count} event{Plural(events.Count)} • {notes} note{Plural(notes)} • {moods} mood{Plural(moods)} • {important} important");
    }

    private static string Plural(int count) => count == 1 ? "" : "s";

    private static int CmdPromise(string[] rest)
    {
        if (rest.Length > 0 && rest[0].Equals("me", StringComparison.OrdinalIgnoreCase))
        {
            PrintProposal();
            return 0;
        }

        Console.WriteLine("usage: pulse promise me");
        return 2;
    }

    private static void PrintProposal()
    {
        Console.WriteLine("Past tomorrow’s worry,");
        Console.WriteLine("and past yesteryear’s take,");
        Console.WriteLine();
        Console.WriteLine("will you make me");
        Console.WriteLine("the happiest man alive?");
        Console.WriteLine();
        Console.WriteLine("Will you make me live?");
        Console.WriteLine("Make me mad?");
        Console.WriteLine();
        Console.WriteLine("Today,");
        Console.WriteLine("tomorrow,");
        Console.WriteLine("and forever");
        Console.WriteLine();
        Console.WriteLine("is what I long to hear.");
        Console.WriteLine();
        Console.WriteLine("Will you be mine,");
        Console.WriteLine("evermore?");
        Console.WriteLine();
        Console.WriteLine("My heart,");
        Console.WriteLine("my love,");
        Console.WriteLine("my dear");
        Console.WriteLine();
        Console.WriteLine("Will you marry me,");
        Console.WriteLine("Quandranea?");
    }


    private static int CmdYear(string[] rest)
    {
        int year = DateTime.Now.Year;

        if (rest.Length > 0)
        {
            if (!int.TryParse(rest[0], out year) || year < 1)
            {
                Console.Error.WriteLine("usage: pulse year [yyyy]");
                Console.Error.WriteLine("example: pulse year 2026");
                return 2;
            }
        }

        var important = ReadEvents()
            .Where(e => TryParseIso(e.Timestamp, out var ts)
                && ts.LocalDateTime.Year == year
                && (e.Type.Equals("important", StringComparison.OrdinalIgnoreCase)
                    || e.Category.Equals("important", StringComparison.OrdinalIgnoreCase)))
            .OrderBy(e => ParseIso(e.Timestamp))
            .ToList();

        Console.WriteLine($"Pulse — {year}");
        Console.WriteLine();

        if (important.Count == 0)
        {
            Console.WriteLine("No important events for this year.");
            return 0;
        }

        int currentMonth = -1;

        foreach (var ev in important)
        {
            var ts = ParseIso(ev.Timestamp).ToLocalTime();

            if (ts.Month != currentMonth)
            {
                currentMonth = ts.Month;
                Console.WriteLine(ts.ToString("MMMM", CultureInfo.InvariantCulture) + ":");
            }

            string date = ts.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            string note = string.IsNullOrWhiteSpace(ev.Note) ? HumanType(ev.Type) : ev.Note;
            string? tags = GetMetaStringArrayDisplay(ev, "tags");
            string tagSuffix = string.IsNullOrWhiteSpace(tags) ? "" : $"  [{tags}]";

            Console.WriteLine($"  {date}  {note}{tagSuffix}");
        }

        Console.WriteLine();
        Console.WriteLine($"{important.Count} important event" + (important.Count == 1 ? "" : "s") + " found.");
        return 0;
    }

    private static void PrintDaySummary(List<PulseEvent> events)
    {
        PrintGroup("Sleep", events.Where(e => e.Category == "sleep").ToList());
        PrintGroup("Meals", events.Where(e => e.Category == "meal").ToList());
        PrintGroup("Body", events.Where(e => e.Category == "body").ToList());
        PrintGroup("Mind", events.Where(e => e.Category == "mental").ToList());
        PrintGroup("Activity", events.Where(e => e.Category == "activity").ToList());
        PrintGroup("Notes", events.Where(e => e.Category is "custom" or "note").ToList());
    }

    private static void PrintGroup(string title, List<PulseEvent> events)
    {
        if (events.Count == 0)
            return;

        Console.WriteLine($"{title}:");

        foreach (var ev in events)
            Console.WriteLine("  " + FormatEventLine(ev).Trim());
    }

    private static ParsedCommand ParseCommand(string command, string[] rest)
    {
        DateTimeOffset ts = DateTimeOffset.Now;
        bool hasExplicitTime = false;
        var rawParts = new List<string>();

        if (rest.Length > 0 && TryParseFlexibleTime(rest[0], out var explicitTs))
        {
            ts = explicitTs;
            hasExplicitTime = true;
            rawParts.AddRange(rest.Skip(1));
        }
        else
        {
            rawParts.AddRange(rest);
        }

        var noteParts = new List<string>();
        var tags = new List<string>();

        foreach (var part in rawParts)
        {
            if (TryParseTag(part, out var tag))
            {
                tags.Add(tag);
                continue;
            }

            noteParts.Add(part);
        }

        string? note = noteParts.Count == 0 ? null : string.Join(' ', noteParts).Trim();

        if (string.IsNullOrWhiteSpace(note))
            note = null;

        return new ParsedCommand
        {
            Command = command,
            Timestamp = ts,
            HasExplicitTime = hasExplicitTime,
            Note = note,
            Tags = tags
        };
    }

    private static bool TryParseDay(string raw, out DateOnly day)
    {
        if (raw.Equals("today", StringComparison.OrdinalIgnoreCase))
        {
            day = DateOnly.FromDateTime(DateTime.Now);
            return true;
        }

        if (raw.Equals("yesterday", StringComparison.OrdinalIgnoreCase))
        {
            day = DateOnly.FromDateTime(DateTime.Now.AddDays(-1));
            return true;
        }

        return DateOnly.TryParseExact(raw, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out day)
            || DateOnly.TryParse(raw, CultureInfo.CurrentCulture, DateTimeStyles.None, out day);
    }

    private static bool TryParseFlexibleTime(string raw, out DateTimeOffset timestamp)
    {
        timestamp = DateTimeOffset.Now;
        string value = raw.Trim();

        if (string.IsNullOrWhiteSpace(value))
            return false;

        var now = DateTimeOffset.Now;

        string normalized = value
            .Replace("a.m.", "am", StringComparison.OrdinalIgnoreCase)
            .Replace("p.m.", "pm", StringComparison.OrdinalIgnoreCase)
            .Replace("AM", "am", StringComparison.Ordinal)
            .Replace("PM", "pm", StringComparison.Ordinal);

        string[] formats =
        {
            "h:mmtt", "htt", "h:mm tt", "h tt",
            "H:mm", "HH:mm",
            "yyyy-MM-ddTHH:mm:sszzz", "yyyy-MM-ddTHH:mmzzz",
            "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm",
            "yyyy-MM-dd h:mmtt", "yyyy-MM-dd h:mm tt",
            "yyyy-MM-dd H:mm", "yyyy-MM-dd HH:mm",
            "M/d/yyyy h:mmtt", "M/d/yyyy h:mm tt",
            "M/d/yyyy H:mm", "M/d/yyyy HH:mm"
        };

        foreach (var fmt in formats)
        {
            if (!DateTimeOffset.TryParseExact(normalized, fmt, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var dto))
                continue;

            if (!fmt.Contains('y'))
            {
                timestamp = new DateTimeOffset(now.Year, now.Month, now.Day, dto.Hour, dto.Minute, dto.Second, now.Offset);
                return true;
            }

            if (dto.Offset == TimeSpan.Zero && !normalized.Contains('+') && !normalized.EndsWith("Z", StringComparison.OrdinalIgnoreCase))
                timestamp = new DateTimeOffset(dto.DateTime, now.Offset);
            else
                timestamp = dto;

            return true;
        }

        if (DateTimeOffset.TryParse(normalized, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed))
        {
            timestamp = parsed;
            return true;
        }

        return false;
    }

    private static PulseEvent NewEvent(string type, string category, DateTimeOffset ts, string? note)
    {
        return new PulseEvent
        {
            Id = NewId(),
            Timestamp = ToIso(ts),
            Type = type,
            Category = category,
            Note = note,
            Source = "cli",
            Meta = new Dictionary<string, JsonElement>()
        };
    }

    private static string NewId()
    {
        string raw = $"{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds():x}-{Guid.NewGuid():N}";
        return raw.Length <= 28 ? raw : raw[..28];
    }

    private static string ToIso(DateTimeOffset ts)
    {
        return ts.ToString("yyyy-MM-dd'T'HH:mm:sszzz", CultureInfo.InvariantCulture);
    }

    private static DateTimeOffset ParseIso(string iso)
    {
        return DateTimeOffset.Parse(iso, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal);
    }

    private static bool TryParseIso(string iso, out DateTimeOffset timestamp)
    {
        return DateTimeOffset.TryParse(iso, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out timestamp);
    }

    private static void AppendEvent(PulseEvent ev, bool fsync = false)
    {
        string line = JsonSerializer.Serialize(ev, JsonOptions) + Environment.NewLine;

        using var stream = new FileStream(EventsPath, FileMode.Append, FileAccess.Write, FileShare.Read);
        byte[] bytes = Encoding.UTF8.GetBytes(line);
        stream.Write(bytes, 0, bytes.Length);

        if (fsync)
            stream.Flush(flushToDisk: true);
    }

    private static IEnumerable<PulseEvent> ReadEvents()
    {
        if (!File.Exists(EventsPath))
            yield break;

        foreach (var line in File.ReadLines(EventsPath))
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            PulseEvent? ev = null;

            try
            {
                ev = JsonSerializer.Deserialize<PulseEvent>(line, JsonOptions);
            }
            catch
            {
                // JSONL survives partial corruption. Bad lines are skipped.
            }

            if (ev is not null)
            {
                ev.Meta ??= new Dictionary<string, JsonElement>();
                yield return ev;
            }
        }
    }

    private static void WriteAllEvents(List<PulseEvent> events)
    {
        string tmp = EventsPath + ".tmp";

        using (var writer = new StreamWriter(tmp, false, Encoding.UTF8))
        {
            foreach (var ev in events)
                writer.WriteLine(JsonSerializer.Serialize(ev, JsonOptions));
        }

        File.Move(tmp, EventsPath, overwrite: true);
    }

    private static PulseState LoadState()
    {
        if (!File.Exists(StatePath))
            return new PulseState();

        try
        {
            return JsonSerializer.Deserialize<PulseState>(File.ReadAllText(StatePath), JsonOptions) ?? new PulseState();
        }
        catch
        {
            return new PulseState();
        }
    }

    private static void SaveState(PulseState state)
    {
        AtomicWriteText(StatePath, JsonSerializer.Serialize(state, PrettyJsonOptions));
    }

    private static PulseConfig LoadConfig()
    {
        EnsureDefaultConfig();

        try
        {
            return JsonSerializer.Deserialize<PulseConfig>(File.ReadAllText(ConfigPath), JsonOptions) ?? new PulseConfig();
        }
        catch
        {
            return new PulseConfig();
        }
    }

    private static void SaveConfig(PulseConfig config)
    {
        AtomicWriteText(ConfigPath, JsonSerializer.Serialize(config, PrettyJsonOptions));
    }

    private static void EnsureDefaultConfig()
    {
        if (!File.Exists(ConfigPath))
            SaveConfig(new PulseConfig());
    }

    private static void AtomicWriteText(string path, string text)
    {
        string tmp = path + ".tmp";
        File.WriteAllText(tmp, text);
        File.Move(tmp, path, overwrite: true);
    }


    private static bool TryParseTag(string part, out string tag)
    {
        tag = "";

        if (!part.StartsWith("tag:", StringComparison.OrdinalIgnoreCase))
            return false;

        tag = part["tag:".Length..].Trim();

        if (string.IsNullOrWhiteSpace(tag))
            return false;

        tag = tag.Trim('"', '\'');
        return !string.IsNullOrWhiteSpace(tag);
    }

    private static void ApplyTags(PulseEvent ev, List<string> tags)
    {
        if (tags.Count == 0)
            return;

        var clean = tags
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (clean.Count > 0)
            SetMeta(ev, "tags", clean);
    }

    private static string? GetMetaStringArrayDisplay(PulseEvent ev, string key)
    {
        if (ev.Meta is null || !ev.Meta.TryGetValue(key, out var el))
            return null;

        if (el.ValueKind != JsonValueKind.Array)
            return null;

        var values = new List<string>();

        foreach (var item in el.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.String)
            {
                var value = item.GetString();
                if (!string.IsNullOrWhiteSpace(value))
                    values.Add(value);
            }
        }

        return values.Count == 0 ? null : string.Join(", ", values);
    }

    private static void SetMeta<T>(PulseEvent ev, string key, T value)
    {
        ev.Meta[key] = JsonSerializer.SerializeToElement(value, JsonOptions);
    }

    private static bool TryGetMetaInt(PulseEvent ev, string key, out int value)
    {
        value = 0;

        if (ev.Meta is null || !ev.Meta.TryGetValue(key, out var el))
            return false;

        if (el.ValueKind == JsonValueKind.Number && el.TryGetInt32(out var n))
        {
            value = n;
            return true;
        }

        if (el.ValueKind == JsonValueKind.String && int.TryParse(el.GetString(), out var s))
        {
            value = s;
            return true;
        }

        return false;
    }

    private static string CategoryForToggle(string cmd)
    {
        return cmd switch
        {
            "meal" or "breakfast" or "lunch" or "dinner" or "snack" => "meal",
            "walk" or "exercise" => "body",
            "work" or "study" => "activity",
            "break" => "rest",
            _ => "activity"
        };
    }

    private static string CategoryForPoint(string cmd)
    {
        return cmd switch
        {
            "wake" => "sleep",
            "pain" => "body",
            "mood" => "mental",
            "energy" => "body",
            "dream" => "mental",
            "journal" => "note",
            "note" => "note",
            "important" => "important",
            _ => "custom"
        };
    }

    private static string Title(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return value;

        return char.ToUpperInvariant(value[0]) + value[1..];
    }

    private static string HumanType(string value)
    {
        return value.Replace('_', ' ');
    }

    private static string FriendlyTime(DateTimeOffset ts)
    {
        return ts.ToLocalTime().ToString("h:mm tt", CultureInfo.InvariantCulture);
    }

    private static string FormatMinutes(int minutes)
    {
        int h = minutes / 60;
        int m = minutes % 60;
        return h > 0 ? $"{h}h {m:00}m" : $"{m}m";
    }

    private static string NoteSuffix(string? note)
    {
        return string.IsNullOrWhiteSpace(note) ? "" : $" — {note}";
    }

    private static string FormatEventLine(PulseEvent ev)
    {
        var ts = ParseIso(ev.Timestamp);
        string note = string.IsNullOrWhiteSpace(ev.Note) ? "" : $" — {ev.Note}";

        string duration = TryGetMetaInt(ev, "duration_minutes", out var mins)
            ? $" ({FormatMinutes(mins)})"
            : "";

        return $"{FriendlyTime(ts),8}  {HumanType(ev.Type),-18} {ev.Category,-8}{duration}{note}";
    }
}