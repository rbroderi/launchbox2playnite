using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Collections.ObjectModel;
using System.Text;
using Playnite.SDK;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace LaunchBoxYamlImporter
{
    public class LaunchBoxYamlImporter : GenericPlugin
    {
        private readonly ILogger logger;
        private readonly string logFilePath;

        // Give this plugin a stable GUID; if you already have one, keep that instead.
        public override Guid Id { get; } = new Guid("7b9f7c07-9f7f-4c8e-a37c-5e28c310aa01");

        public LaunchBoxYamlImporter(IPlayniteAPI api) : base(api)
        {
            logger = LogManager.GetLogger();
            logger.Info("LaunchBoxYamlImporter initialized");
            logFilePath = GetLogFilePath();
            AppendLog("LaunchBoxYamlImporter initialized (ctor)");

            Properties = new GenericPluginProperties
            {
                HasSettings = false
            };
        }

        public override IEnumerable<MainMenuItem> GetMainMenuItems(GetMainMenuItemsArgs args)
        {
            yield return new MainMenuItem
            {
                Description = "Import LaunchBox YAML…",
                MenuSection = "@LaunchBox YAML",
                Action = _ => ImportFromYaml()
            };
        }

        private void ImportFromYaml()
        {
            AppendLog("ImportFromYaml started");
            // Uses IDialogsFactory.SelectFile – this exists in SDK 6.14.0
            var path = PlayniteApi.Dialogs.SelectFile(
                "YAML files (*.yaml;*.yml)|*.yaml;*.yml|All files (*.*)|*.*");

            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                var yamlText = File.ReadAllText(path);

                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(NullNamingConvention.Instance)
                    .IgnoreUnmatchedProperties()
                    .Build();

                List<LaunchBoxGameYaml> games;

                // Case 1: file is just a list: "- title: ..."
                try
                {
                    games = deserializer.Deserialize<List<LaunchBoxGameYaml>>(yamlText)
                             ?? new List<LaunchBoxGameYaml>();
                }
                catch
                {
                    // Case 2: file has a root "games:" key
                    var root = deserializer.Deserialize<LaunchBoxRootYaml>(yamlText);
                    games = root?.Games ?? new List<LaunchBoxGameYaml>();
                }

                if (games.Count == 0)
                {
                    PlayniteApi.Dialogs.ShowMessage(
                        "No games found in YAML file.",
                        "LaunchBox YAML Import");
                    return;
                }

                var yamlBasePath = Path.GetDirectoryName(path) ?? Directory.GetCurrentDirectory();

                var stats = ImportGames(games, yamlBasePath, out var mediaIssues);

                AppendLog($"ImportFromYaml completed: {stats.Processed} games processed. {stats}");
                PlayniteApi.Dialogs.ShowMessage(
                    $"Imported {stats.Added} new and updated {stats.Updated} games (processed {stats.Processed}/{stats.Total}).",
                    "LaunchBox YAML Import");

                if (mediaIssues.Count > 0)
                {
                    var details = string.Join(Environment.NewLine, mediaIssues.Take(10));
                    var warningText = $"Imported with {mediaIssues.Count} media issues. Showing up to 10:{Environment.NewLine}{details}";
                    AppendLog(warningText);
                    PlayniteApi.Dialogs.ShowMessage(
                        warningText,
                        "LaunchBox YAML Import");
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Import failed: {ex.Message}");
                PlayniteApi.Dialogs.ShowErrorMessage(
                    ex.Message,
                    "LaunchBox YAML Import Error");
            }
        }

        private ImportStats ImportGames(List<LaunchBoxGameYaml> sourceGames, string basePath, out List<string> mediaIssues)
        {
            var db = PlayniteApi.Database;
            mediaIssues = new List<string>();

            var coverSources = sourceGames.Count(g => !string.IsNullOrWhiteSpace(g.Image));
            var bgSources = sourceGames.Count(g => !string.IsNullOrWhiteSpace(g.BackgroundImage));
            var iconSources = sourceGames.Count(g => !string.IsNullOrWhiteSpace(g.Icon));
            var manualSources = sourceGames.Count(g => !string.IsNullOrWhiteSpace(g.Manual));
            AppendLog($"Deserialized media counts: cover={coverSources}, background={bgSources}, icon={iconSources}, manual={manualSources}");

            // Cache platforms by name (case-insensitive)
            var platformsByName = db.Platforms
                .ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

            var gamesByLaunchBoxId = new Dictionary<string, Game>(StringComparer.OrdinalIgnoreCase);
            var gamesByTitle = new Dictionary<string, Game>(StringComparer.OrdinalIgnoreCase);

            foreach (var existingGame in db.Games)
            {
                if (!string.IsNullOrWhiteSpace(existingGame.GameId))
                {
                    gamesByLaunchBoxId[existingGame.GameId] = existingGame;
                }

                if (!string.IsNullOrWhiteSpace(existingGame.Name))
                {
                    gamesByTitle[existingGame.Name] = existingGame;
                }
            }

            var stats = new ImportStats
            {
                Total = sourceGames.Count
            };

            var seenLaunchBoxIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenTitles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (db.BufferedUpdate())
            {
                foreach (var src in sourceGames)
                {
                    var title = !string.IsNullOrWhiteSpace(src.Title)
                        ? src.Title
                        : src.Name;

                    if (string.IsNullOrWhiteSpace(title))
                    {
                        stats.SkippedNoTitle++;
                        continue;
                    }

                    var titleKey = title!;

                    if (!seenTitles.Add(titleKey))
                    {
                        stats.DuplicateTitles++;
                    }

                    var launchBoxIdRaw = src.LaunchBoxId?.Trim();
                    var launchBoxId = string.IsNullOrWhiteSpace(launchBoxIdRaw) ? null : launchBoxIdRaw;
                    Game? existing = null;

                    if (launchBoxId != null &&
                        gamesByLaunchBoxId.TryGetValue(launchBoxId, out var byId))
                    {
                        existing = byId;
                        stats.MatchedById++;
                    }
                    else if (gamesByTitle.TryGetValue(titleKey, out var byName))
                    {
                        existing = byName;
                        stats.MatchedByTitle++;
                    }

                    var game = existing ?? new Game(titleKey)
                    {
                        Id = existing?.Id ?? Guid.NewGuid()
                    };

                    if (launchBoxId != null)
                    {
                        if (!seenLaunchBoxIds.Add(launchBoxId))
                        {
                            stats.DuplicateLaunchBoxIds++;
                        }

                        game.GameId = launchBoxId;
                        gamesByLaunchBoxId[launchBoxId] = game;
                    }

                    gamesByTitle[titleKey] = game;

                    var sortName = !string.IsNullOrWhiteSpace(src.SortTitle)
                        ? src.SortTitle
                        : (!string.IsNullOrWhiteSpace(src.SortingName)
                            ? src.SortingName
                            : titleKey);

                    // Basic text fields
                    game.SortingName = sortName;

                    // If you had "description" separate from "notes" you can map accordingly
                    game.Description = src.Description ?? src.Notes;
                    game.Notes = src.Notes ?? src.Description;

                    // Favorite flag from LaunchBox
                    game.Favorite = src.Favorite;

                    // Mark as installed (these are local batch/exe launchers)
                    game.IsInstalled = true;

                    var playPathSource = src.PlayAction?.Path ?? src.ApplicationPath;
                    var playArgs = src.PlayAction?.Arguments ?? src.CommandLine;
                    var workingDirSource = src.PlayAction?.WorkingDir;

                    var resolvedPlayPath = ResolvePath(playPathSource, basePath);
                    var resolvedWorkingDir = ResolvePath(workingDirSource, basePath);
                    var resolvedConfigPath = ResolvePath(src.ConfigurationPath, basePath);
                    var resolvedRootFolder = ResolvePath(src.RootFolder, basePath);

                    // Install directory from RootFolder when available, else derive from play path
                    if (!string.IsNullOrWhiteSpace(resolvedRootFolder) && Directory.Exists(resolvedRootFolder))
                    {
                        game.InstallDirectory = resolvedRootFolder;
                    }
                    else
                    {
                        game.InstallDirectory = GetInstallDirectory(resolvedPlayPath);
                    }

                    // Replace Play action based on ApplicationPath / CommandLine
                    game.GameActions = BuildGameActions(
                        resolvedPlayPath,
                        playArgs,
                        resolvedWorkingDir,
                        resolvedConfigPath);

                    // Platform mapping – uses PlatformIds, which exists in 6.14
                    if (!string.IsNullOrWhiteSpace(src.Platform))
                    {
                        var platformName = src.Platform!;

                        if (!platformsByName.TryGetValue(platformName, out var platform))
                        {
                            platform = new Platform(platformName);
                            db.Platforms.Add(platform);
                            platformsByName[platform.Name] = platform;
                        }

                        game.PlatformIds = new List<Guid> { platform.Id };
                    }

                    var isNewGame = existing == null;
                    if (isNewGame)
                    {
                        db.Games.Add(game);
                        stats.Added++;
                    }
                    else
                    {
                        stats.Updated++;
                    }

                    LinkMedia(game, src, basePath, mediaIssues);
                    db.Games.Update(game);
                }
            }

            AppendLog($"Import stats: {stats}");
            return stats;
        }

        private static string GetInstallDirectory(string? sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                return string.Empty;
            }

            try
            {
                var full = Path.GetFullPath(sourcePath);
                return Path.GetDirectoryName(full) ?? string.Empty;
            }
            catch
            {
                // Worst case, leave it empty; the play action will still have the path.
                return string.Empty;
            }
        }

        private static ObservableCollection<GameAction> BuildGameActions(
            string? applicationPath,
            string? commandLine,
            string? workingDirectory,
            string? configurationPath)
        {
            var actions = new ObservableCollection<GameAction>();

            if (string.IsNullOrWhiteSpace(applicationPath))
            {
                return actions;
            }

            var workingDir = !string.IsNullOrWhiteSpace(workingDirectory)
                ? workingDirectory!
                : SafeDirName(applicationPath!);

            var act = new GameAction
            {
                Name = "Play",
                Path = applicationPath,
                Arguments = commandLine ?? string.Empty,
                WorkingDir = workingDir,
                Type = GameActionType.File,
                IsPlayAction = true
            };

            actions.Add(act);

            if (!string.IsNullOrWhiteSpace(configurationPath))
            {
                var configPath = configurationPath!;
                var configWorkingDir = SafeDirName(configPath);

                actions.Add(new GameAction
                {
                    Name = "Install / Configure",
                    Path = configPath,
                    Arguments = string.Empty,
                    WorkingDir = configWorkingDir,
                    Type = GameActionType.File,
                    IsPlayAction = false
                });
            }
            return actions;
        }

        private void LinkMedia(Game game, LaunchBoxGameYaml src, string basePath, List<string> mediaIssues)
        {
            TryAssignMedia("cover", id => game.CoverImage = id, g => g.CoverImage, src.Image, basePath, game, mediaIssues);
            TryAssignMedia("background", id => game.BackgroundImage = id, g => g.BackgroundImage, src.BackgroundImage, basePath, game, mediaIssues);
            TryAssignMedia("icon", id => game.Icon = id, g => g.Icon, src.Icon, basePath, game, mediaIssues);
            AssignManualPath(game, src.Manual, basePath, mediaIssues);
        }

        private void TryAssignMedia(
            string mediaLabel,
            Action<string> setter,
            Func<Game, string?> currentGetter,
            string? sourcePath,
            string basePath,
            Game game,
            List<string> mediaIssues)
        {
            AppendLog($"Game '{game.Name}': processing {mediaLabel}");
            var resolved = ResolvePath(sourcePath, basePath);
            AppendLog($"Game '{game.Name}': {mediaLabel} resolved path '{resolved}'");

            if (string.IsNullOrWhiteSpace(resolved))
            {
                AppendLog($"Game '{game.Name}': {mediaLabel} source path empty");
                return;
            }

            if (!File.Exists(resolved))
            {
                if (!string.IsNullOrWhiteSpace(sourcePath))
                {
                    var msg = $"Missing {mediaLabel} for {game.Name}: {resolved}";
                    mediaIssues.Add(msg);
                    AppendLog(msg);
                }
                return;
            }

            var mediaId = ImportFileForGame(resolved, game, mediaLabel, mediaIssues);
            if (string.IsNullOrWhiteSpace(mediaId))
            {
                AppendLog($"Game '{game.Name}': {mediaLabel} import returned empty ID");
                return;
            }

            var existingId = currentGetter(game);
            if (!string.IsNullOrWhiteSpace(existingId))
            {
                PlayniteApi.Database.RemoveFile(existingId);
            }

            AppendLog($"Game '{game.Name}': {mediaLabel} assigned file ID {mediaId}");
            setter(mediaId!);
        }

        private void AssignManualPath(Game game, string? sourcePath, string basePath, List<string> mediaIssues)
        {
            var resolved = ResolvePath(sourcePath, basePath);
            AppendLog($"Game '{game.Name}': manual resolved path '{resolved}'");

            if (string.IsNullOrWhiteSpace(resolved))
            {
                AppendLog($"Game '{game.Name}': manual source path empty");
                return;
            }

            if (!File.Exists(resolved))
            {
                var msg = $"Missing manual for {game.Name}: {resolved}";
                mediaIssues.Add(msg);
                AppendLog(msg);
                return;
            }

            try
            {
                resolved = Path.GetFullPath(resolved);
                resolved = resolved.Replace('/', '\\');
            }
            catch
            {
                // keep original string if we can't expand
            }

            if (string.Equals(game.Manual, resolved, StringComparison.OrdinalIgnoreCase))
            {
                AppendLog($"Game '{game.Name}': manual path unchanged");
                return;
            }

            game.Manual = resolved;
            AppendLog($"Game '{game.Name}': manual path set to absolute path");
        }


        private string? ImportFileForGame(string filePath, Game game, string mediaLabel, List<string> mediaIssues)
        {
            try
            {
                if (game.Id == Guid.Empty)
                {
                    var missingIdMsg = $"Cannot import {mediaLabel} for {game.Name}: game has no ID.";
                    mediaIssues.Add(missingIdMsg);
                    AppendLog(missingIdMsg);
                    return null;
                }

                var dbId = PlayniteApi.Database.AddFile(filePath, game.Id);
                AppendLog($"Game '{game.Name}': {mediaLabel} added to DB with id {dbId}");
                return dbId;
            }
            catch (Exception ex)
            {
                var msg = $"Failed to import {mediaLabel} for {game.Name} from {filePath}: {ex.Message}";
                mediaIssues.Add(msg);
                AppendLog(msg);
                return null;
            }
        }

        private void AppendLog(string message)
        {
            try
            {
                var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}";
                File.AppendAllLines(logFilePath, new[] { line }, Encoding.UTF8);
            }
            catch
            {
                // swallow logging exceptions
            }
        }

        private string GetLogFilePath()
        {
            var baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Playnite",
                "ExtensionsData",
                Id.ToString());

            Directory.CreateDirectory(baseDir);
            return Path.Combine(baseDir, "launchbox2playnite.log");
        }

        private static string ResolvePath(string? sourcePath, string basePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                return string.Empty;
            }

            var relativePath = sourcePath!;

            if (Path.IsPathRooted(relativePath))
            {
                return relativePath;
            }

            try
            {
                return Path.GetFullPath(Path.Combine(basePath, relativePath));
            }
            catch
            {
                return relativePath;
            }
        }

        private static string SafeDirName(string path)
        {
            try
            {
                return Path.GetDirectoryName(path) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private sealed class ImportStats
        {
            public int Total { get; set; }
            public int Added { get; set; }
            public int Updated { get; set; }
            public int SkippedNoTitle { get; set; }
            public int MatchedById { get; set; }
            public int MatchedByTitle { get; set; }
            public int DuplicateLaunchBoxIds { get; set; }
            public int DuplicateTitles { get; set; }

            public int Processed => Added + Updated;

            public override string ToString()
            {
                return $"total={Total}, added={Added}, updated={Updated}, matchedById={MatchedById}, matchedByTitle={MatchedByTitle}, skippedNoTitle={SkippedNoTitle}, duplicateLaunchBoxIds={DuplicateLaunchBoxIds}, duplicateTitles={DuplicateTitles}";
            }
        }

        // --- YAML DTOs ---

        // Root wrapper if you export as:
        // games:
        //   - title: ...
        private sealed class LaunchBoxRootYaml
        {
            public List<LaunchBoxGameYaml> Games { get; set; } = new List<LaunchBoxGameYaml>();
        }

        // Per-game record. Property names are in camelCase to
        // match YamlDotNet’s CamelCaseNamingConvention.
        private sealed class LaunchBoxGameYaml
        {
            public string? Id { get; set; }
            public string? Title { get; set; }
            public string? Name { get; set; }
            public string? SortTitle { get; set; }
            public string? SortingName { get; set; }

            public string? Platform { get; set; }

            public string? Image { get; set; }
            public string? BackgroundImage { get; set; }
            public string? Icon { get; set; }
            public List<string>? Screenshots { get; set; }
            public List<string>? Videos { get; set; }

            public string? ApplicationPath { get; set; }
            public string? CommandLine { get; set; }
            public string? ConfigurationPath { get; set; }
            public string? RootFolder { get; set; }

            public PlayActionYaml? PlayAction { get; set; }
            public List<RomYaml>? Roms { get; set; }

            public string? Manual { get; set; }
            public string? LaunchBoxId { get; set; }

            public string? Description { get; set; }
            public string? Notes { get; set; }

            public bool Favorite { get; set; }
        }

        private sealed class PlayActionYaml
        {
            public string? Path { get; set; }
            public string? WorkingDir { get; set; }
            public string? Arguments { get; set; }
        }

        private sealed class RomYaml
        {
            public string? Path { get; set; }
            public string? Size { get; set; }
        }
    }
}
