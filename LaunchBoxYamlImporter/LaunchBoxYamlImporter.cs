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
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
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

                var imported = ImportGames(games);

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

        private int ImportGames(List<LaunchBoxGameYaml> sourceGames)
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
                    if (string.IsNullOrWhiteSpace(src.Title))
                    {
                        continue;
                    }

                    // Very simple match: by name only.
                    // If you later add a stable ID from LaunchBox, you can match on that.
                    var existing = db.Games.FirstOrDefault(g =>
                        string.Equals(g.Name, src.Title, StringComparison.OrdinalIgnoreCase));

                    var game = existing ?? new Game(src.Title);

                    // Basic text fields
                    game.SortingName = !string.IsNullOrWhiteSpace(src.SortTitle)
                        ? src.SortTitle
                        : src.Title;

                    // If you had "description" separate from "notes" you can map accordingly
                    game.Description = src.Description ?? src.Notes;
                    game.Notes = src.Notes;

                    // Favorite flag from LaunchBox
                    game.Favorite = src.Favorite;

                    // Mark as installed (these are local batch/exe launchers)
                    game.IsInstalled = true;

                    // Install directory from ApplicationPath if we can resolve it
                    game.InstallDirectory = GetInstallDirectory(src.ApplicationPath);

                    // Replace Play action based on ApplicationPath / CommandLine
                    game.GameActions = BuildGameActions(src.ApplicationPath, src.CommandLine);

                    // Platform mapping – uses PlatformIds, which exists in 6.14
                    if (!string.IsNullOrWhiteSpace(src.Platform))
                    {
                        if (!platformsByName.TryGetValue(src.Platform, out var platform))
                        {
                            platform = new Platform(src.Platform);
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

        private static string GetInstallDirectory(string? applicationPath)
        {
            if (string.IsNullOrWhiteSpace(applicationPath))
            {
                return string.Empty;
            }

            try
            {
                // If it's relative, this will resolve relative to current working directory.
                var full = Path.GetFullPath(applicationPath);
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
            string? commandLine)
        {
            var actions = new ObservableCollection<GameAction>();

            if (string.IsNullOrWhiteSpace(applicationPath))
            {
                return actions;
            }

            var act = new GameAction
            {
                Name = "Play",
                Path = applicationPath,
                Arguments = commandLine ?? string.Empty,
                WorkingDir = SafeDirName(applicationPath),
                Type = GameActionType.File,
                IsPlayAction = true
            };

            actions.Add(act);
            return actions;
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
            public string? SortTitle { get; set; }

            public string? Platform { get; set; }

            public string? ApplicationPath { get; set; }
            public string? CommandLine { get; set; }

            public string? Description { get; set; }
            public string? Notes { get; set; }

            public bool Favorite { get; set; }
        }
    }
}
