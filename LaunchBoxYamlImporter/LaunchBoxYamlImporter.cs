using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Collections.ObjectModel;
using Playnite.SDK;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace LaunchBoxYamlImporter
{
    public class LaunchBoxYamlImporter : GenericPlugin
    {
        private static readonly ILogger logger = LogManager.GetLogger();

        // Give this plugin a stable GUID; if you already have one, keep that instead.
        public override Guid Id { get; } = new Guid("7b9f7c07-9f7f-4c8e-a37c-5e28c310aa01");

        public LaunchBoxYamlImporter(IPlayniteAPI api) : base(api)
        {
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

                var imported = ImportGames(games, yamlBasePath);

                PlayniteApi.Dialogs.ShowMessage(
                    $"Imported or updated {imported} games from LaunchBox YAML.",
                    "LaunchBox YAML Import");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to import LaunchBox YAML.");
                PlayniteApi.Dialogs.ShowErrorMessage(
                    ex.Message,
                    "LaunchBox YAML Import Error");
            }
        }

        private int ImportGames(List<LaunchBoxGameYaml> sourceGames, string basePath)
        {
            var db = PlayniteApi.Database;

            // Cache platforms by name (case-insensitive)
            var platformsByName = db.Platforms
                .ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

            var imported = 0;

            using (db.BufferedUpdate())
            {
                foreach (var src in sourceGames)
                {
                    var title = !string.IsNullOrWhiteSpace(src.Title)
                        ? src.Title
                        : src.Name;

                    if (string.IsNullOrWhiteSpace(title))
                    {
                        continue;
                    }

                    // Very simple match: by name only.
                    // If you later add a stable ID from LaunchBox, you can match on that.
                    var existing = db.Games.FirstOrDefault(g =>
                        string.Equals(g.Name, title, StringComparison.OrdinalIgnoreCase));

                    var game = existing ?? new Game(title);

                    var sortName = !string.IsNullOrWhiteSpace(src.SortTitle)
                        ? src.SortTitle
                        : (!string.IsNullOrWhiteSpace(src.SortingName)
                            ? src.SortingName
                            : title);

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

                    LinkMedia(game, src, basePath);

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

                    if (existing == null)
                    {
                        db.Games.Add(game);
                    }
                    else
                    {
                        db.Games.Update(game);
                    }

                    imported++;
                }
            }

            return imported;
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

        private void LinkMedia(Game game, LaunchBoxGameYaml src, string basePath)
        {
            TryAssignImage(path => game.CoverImage = path, src.Image, basePath);
            TryAssignImage(path => game.BackgroundImage = path, src.BackgroundImage, basePath);
            TryAssignImage(path => game.Icon = path, src.Icon, basePath);

            var manualPath = ResolvePath(src.Manual, basePath);
            if (!string.IsNullOrWhiteSpace(manualPath) && File.Exists(manualPath))
            {
                game.Manual = manualPath;
            }
        }

        private void TryAssignImage(Action<string> setter, string? sourcePath, string basePath)
        {
            var resolved = ResolvePath(sourcePath, basePath);
            if (!string.IsNullOrWhiteSpace(resolved) && File.Exists(resolved))
            {
                setter(resolved);
            }
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
