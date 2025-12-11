using System;
using System.Linq;
using System.Reflection;
using System.IO;
using LiteDB;

var sdkPath = @"C:\Users\richa\AppData\Local\Playnite\Playnite.SDK.dll";
var asm = Assembly.LoadFrom(sdkPath);
var gameType = asm.GetType("Playnite.SDK.Models.Game");

Console.WriteLine($"Properties on {gameType?.FullName}");
var props = gameType?.GetProperties()?.OrderBy(p => p.Name).ToList() ?? new List<PropertyInfo>();
foreach (var prop in props)
{
	Console.WriteLine($"{prop.Name} : {prop.PropertyType}");
}

var ctor = gameType?.GetConstructor(Type.EmptyTypes);
var instance = ctor?.Invoke(null);
var idProp = gameType?.GetProperty("Id");
Console.WriteLine($"New Game Id: {idProp?.GetValue(instance)}");

var dbPath = @"C:\Users\richa\AppData\Roaming\Playnite\library\games.db";
using var db = new LiteDatabase(new ConnectionString
{
	Filename = dbPath,
	ReadOnly = true,
	Connection = ConnectionType.Direct
});

var games = db.GetCollection("games");
var sample = games.FindOne(Query.All());
if (sample is null)
{
	Console.WriteLine("No games in database.");
}
else
{
	Console.WriteLine($"Sample game: {sample["Name"]}");
	if (sample.TryGetValue("CoverImage", out var cover))
	{
		Console.WriteLine($"CoverImage field: {cover}");
	}
}
