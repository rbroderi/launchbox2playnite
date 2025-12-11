using System;
using System.Collections.Generic;
using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

var file = args.Length > 0 ? args[0] : "../playnite_import_games.yaml";

if (!File.Exists(file))
{
	Console.WriteLine($"File not found: {Path.GetFullPath(file)}");
	return;
}

var yamlText = File.ReadAllText(file);

List<LaunchBoxGameYaml>? games = null;

var camel = new DeserializerBuilder()
	.WithNamingConvention(CamelCaseNamingConvention.Instance)
	.IgnoreUnmatchedProperties()
	.Build();

try
{
	games = camel.Deserialize<List<LaunchBoxGameYaml>>(yamlText);
	Console.WriteLine($"CamelCase count: {games?.Count ?? 0}");
}
catch (Exception ex)
{
	Console.WriteLine($"CamelCase failed: {ex.Message}");
}

var exact = new DeserializerBuilder()
	.WithNamingConvention(NullNamingConvention.Instance)
	.IgnoreUnmatchedProperties()
	.Build();

try
{
	games = exact.Deserialize<List<LaunchBoxGameYaml>>(yamlText);
	Console.WriteLine($"NullNaming count: {games?.Count ?? 0}");
	if (games != null)
	{
		var imageCount = games.Count(g => !string.IsNullOrWhiteSpace(g.Image));
		Console.WriteLine($"Games with Image: {imageCount}");
		var bgCount = games.Count(g => !string.IsNullOrWhiteSpace(g.BackgroundImage));
		var iconCount = games.Count(g => !string.IsNullOrWhiteSpace(g.Icon));
		var manualCount = games.Count(g => !string.IsNullOrWhiteSpace(g.Manual));
		Console.WriteLine($"Backgrounds={bgCount} Icons={iconCount} Manuals={manualCount}");
		var first = games.Count > 0 ? games[0] : null;
		if (first != null)
		{
			Console.WriteLine($"First entry title='{first.Title}' name='{first.Name}' platform='{first.Platform}' image='{first.Image}'");
		}

		var starWars = games.FirstOrDefault(g => string.Equals(g.Name, "Star Wars Chess", StringComparison.OrdinalIgnoreCase));
		if (starWars != null)
		{
			Console.WriteLine($"Star Wars Chess image='{starWars.Image}' background='{starWars.BackgroundImage}' icon='{starWars.Icon}' manual='{starWars.Manual}'");
		}
	}
}
catch (Exception ex)
{
	Console.WriteLine($"NullNaming failed: {ex.Message}");
}

public class LaunchBoxGameYaml
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

public class PlayActionYaml
{
	public string? Path { get; set; }
	public string? WorkingDir { get; set; }
	public string? Arguments { get; set; }
}

public class RomYaml
{
	public string? Path { get; set; }
	public string? Size { get; set; }
}
