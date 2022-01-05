using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Linq;
using System.Net.Http;
using System.Net;

namespace FpaExcelResultsParser
{
	using PlayerDataCollection = Dictionary<string, Tuple<string, string>>;
	using InternetPlayerDataCollection = Dictionary<string, PlayerData>;
	using InternetEventSummaryCollection = Dictionary<string, EventSummary>;

	class EventSummary
	{
		public string key { get; set; }
		public string eventName { get; set; }
		public string startDate { get; set; }
		public string endDate { get; set; }

		public override bool Equals(object obj)
		{
			EventSummary other = obj as EventSummary;
			if (other == null)
			{
				return false;
			}

			if (eventName == other.eventName && startDate == other.startDate && endDate == other.endDate)
			{
				return true;
			}

			return false;
		}

		public override int GetHashCode()
		{
			return (key + eventName + startDate + endDate).GetHashCode();
		}
	}

	class TeamData
	{
		public int place = 0;
		public double points = 0;
		public List<string> playerIds = new List<string>();
	}

	class PlayerData
	{
		public string key { get; set; }
		public long lastActive { get; set; }
		public int memebership { get; set; }
		public string lastName { get; set; }
		public long createdAt { get; set; }
		public string country { get; set; }
		public string firstName { get; set; }
		public string gender { get; set; }

		public string FullName
		{
			get { return firstName + " " + lastName; }
		}
	}

	class PlayerDataResponse
	{
		public InternetPlayerDataCollection players { get; set; }
	}

	class EventSummaryResponse
	{
		public InternetEventSummaryCollection allEventSummaryData { get; set; }
	}

	class Program
	{
		private static HttpClient client = null;
		static int errorCount = 0;
		static List<Tuple<string, string>> rawPlayerList = new List<Tuple<string, string>>();
		static PlayerDataCollection playerNameData;
		static InternetPlayerDataCollection internetPlayerData;
		static InternetEventSummaryCollection internetEventSummaries;
		static Dictionary<string, string> rawNameToId = new Dictionary<string, string>();
		static PlayerDataCollection missingPlayers = new PlayerDataCollection();
		static Dictionary<string, int> roundDefinitions = new Dictionary<string, int>()
		{
			{ "Finals", 1 },
			{ "Final", 1 },
			{ "Semifinals", 2 },
			{ "Semifinal", 2 },
			{ "Quarterfinals", 3 },
			{ "Quarterfinal", 3 },
			{ "Preliminary", 4 }
		};
		static HashSet<string> divisionDefinitions = new HashSet<string>()
		{
			"Open Pairs",
			"Open Coop",
			"Open Co-op",
			"Women Pairs",
			"Mixed Pairs",
			"Random Open",
			"Open"
		};
		static List<string> outputMarkups = new List<string>();

		static void Main(string[] args)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			LoadNameData();
			//AddMissingPlayers();

			LoadEvents();

			List<EventSummary> summeries = new List<EventSummary>();

			var files = Directory.GetFiles(@"C:\GitHub\FpaExcelResultsParser\results\");
			foreach (var filename in files)
			{
				if (filename.EndsWith(".xlsx") && !filename.Contains("~"))
				{
					using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
					{
						ExcelWorksheet sheet = FindSheet(package);
						if (sheet != null)
						{
							//ParsePlayers(sheet, filename);

							var summaryData = ParseEventSummary(sheet, filename);
							//Console.WriteLine(JsonSerializer.Serialize(summaryData));
							var internetSummary = internetEventSummaries.Values.FirstOrDefault(x => x.Equals(summaryData));
							if (internetSummary != null)
							{
								summaryData.key = internetSummary.key;
							}
							else
							{
								Console.WriteLine($"Error: Can't find internet event summary for '{filename}'");
							}

							if (!summeries.Contains(summaryData))
							{
								summeries.Add(summaryData);
							}

							ParseResults(sheet, filename, summaryData);

							if (summaryData.eventName.Length == 0)
							{
								Console.WriteLine($"Error: Can't find event name in '{filename}'");
							}
						}
						else
						{
							Console.WriteLine($"Error: Can't find results sheet in '{filename}'");
						}
					}
				}
			}

			AddEvents(summeries);

			WriteResultOutput();

			Console.WriteLine($"Total events: {summeries.Count}  Errors: {errorCount}  Players: {rawPlayerList.Count} Markups: {outputMarkups.Count}");

			//PrintNameData();

			//using (var package = new ExcelPackage(new FileInfo(@"C:\GitHub\FpaExcelResultsParser\results\130105-V01-1302-In_den_Hallen.xlsx")))
			//{
			//	ExcelWorksheet sheet = package.Workbook.Worksheets[1];
			//	var summaryData = ParseEventSummary(sheet);
			//	Console.WriteLine(JsonSerializer.Serialize(summaryData));
			//}
		}

		static void WriteResultOutput()
		{
			int fileIndex = 0;
			foreach (var markup in outputMarkups)
			{
				string path = @"C:\GitHub\FpaExcelResultsParser\output\";
				using (StreamWriter stream = new StreamWriter(path + fileIndex + ".txt"))
				{
					stream.WriteLine(markup);
				}
				++fileIndex;
			}
		}

		static void ParseResults(ExcelWorksheet sheet, string filename, EventSummary summaryData)
		{
			int playerColumn = -1;
			for (int column = 3; column < 10; ++column)
			{
				int row = SearchForPlayers(sheet, column);
				if (row >= 0)
				{
					playerColumn = column;
					break;
				}
			}

			int lastRow = sheet.Dimension.End.Row;
			List<int> playerTagRows = new List<int>();
			for (int row = 0; row <= lastRow; ++row)
			{
				string text = sheet.GetValue(row, playerColumn) as string;
				if (text == "Player")
				{
					playerTagRows.Add(row);
				}
			}
			playerTagRows.Add(lastRow);

			bool isFirstPool = true;
			string lastRound = null;
			StringBuilder output = new StringBuilder();
			for (int i = 1; i < playerTagRows.Count; ++i)
			{
				int startRow = playerTagRows[i - 1];
				int endRow = playerTagRows[i] - 1;
				string division = ParseDivision(sheet, startRow - 1);
				if (division == null || division.Length == 0)
				{
					//Console.WriteLine(filename);

					continue;
				}
				string round = ParseRound(sheet, startRow - 1);
				if (round == null || round.Length == 0)
				{
					Console.WriteLine(filename);
				}

				int placeColumn = FindColumn(sheet, startRow, "Place");
				int totalColumn = FindColumn(sheet, startRow, "Total");
				if (placeColumn < 0 || totalColumn < 0)
				{
					Console.WriteLine(filename);
				}

				if (isFirstPool)
				{
					isFirstPool = false;
					output.AppendLine($"start pools {summaryData.key} {GetDivisionMarkup(division)}");
				}

				int roundNumber = roundDefinitions[round];
				if (lastRound != round)
				{
					lastRound = round;
					output.AppendLine($"round {roundNumber}");
				}

				int place = 0;
				int lastPlace = 0;
				TeamData lastTeamData = null;
				List<TeamData> teams = new List<TeamData>();
				for (int row = startRow; row <= endRow; ++row)
				{
					string rawPlayer = sheet.GetValue(row, playerColumn) as string;
					if (rawPlayer != null && rawNameToId.ContainsKey(rawPlayer))
					{
						try
						{
							int parsePlace = sheet.GetValue<int>(row, placeColumn);
							place = parsePlace > 0 ? parsePlace : place;
						}
						catch
						{
							string parsePlace = sheet.GetValue(row, placeColumn) as string;
							place = parsePlace != null && int.TryParse(GetOnlyNumbers(parsePlace), out var newPlace) ? newPlace : place;
						}

						double total = sheet.GetValue<double>(row, totalColumn);

						if (place < 1)
						{
							Console.WriteLine(rawPlayer + " - " + filename);
						}

						if (lastPlace != place)
						{
							lastPlace = place;
							lastTeamData = new TeamData();
							teams.Add(lastTeamData);
						}

						lastTeamData.place = place;
						lastTeamData.points = total;
						lastTeamData.playerIds.Add(rawNameToId[rawPlayer]);
					}
				}

				string pool = FindColumnContains(sheet, startRow - 1, "Pool") ?? "pool A";
				if (pool == null || pool.Length == 0)
				{
					Console.WriteLine($"Error: Can't find pool in '{filename}'");
				}

				output.AppendLine(pool.Replace("Pool", "pool"));
				foreach (var team in teams)
				{
					output.AppendLine($"{team.place} {string.Join(" ", team.playerIds)} {team.points}");
				}
			}

			output.AppendLine("end");

			outputMarkups.Add(output.ToString());
		}

		static string GetDivisionMarkup(string division)
		{
			return division.Contains(' ') ? $"\"{division}\"" : division;
		}

		private static string GetOnlyNumbers(string input)
		{
			return new string(input.Where(c => char.IsDigit(c)).ToArray());
		}

		static string ParseRound(ExcelWorksheet sheet, int row)
		{
			for (int column = 0; column < 10; ++column)
			{
				string round = sheet.GetValue(row, column) as string;
				if (round != null && roundDefinitions.Keys.Contains(round))
				{
					return round;
				}
				//else if (round != null && round.Length > 0)
				//{
				//	Console.WriteLine(round);
				//}
			}

			return null;
		}

		static string ParseDivision(ExcelWorksheet sheet, int row)
		{
			for (int column = 0; column < 10; ++column)
			{
				string division = sheet.GetValue(row, column) as string;
				if (divisionDefinitions.Contains(division))
				{
					return division;
				}
				//else if (division != null && division.Length > 0)
				//{
				//	Console.WriteLine(division);
				//}
			}

			return null;
		}

		static int FindColumn(ExcelWorksheet sheet, int row, string columnText)
		{
			for (int column = 0; column < 30; ++column)
			{
				string text = sheet.GetValue(row, column) as string;
				if (columnText == text)
				{
					return column;
				}
			}

			return -1;
		}

		static string FindColumnContains(ExcelWorksheet sheet, int row, string columnText)
		{
			for (int column = 0; column < 30; ++column)
			{
				string text = sheet.GetValue(row, column) as string;
				if (text != null && text.Contains(columnText))
				{
					return text;
				}
			}

			return null;
		}

		static void AddMissingPlayers()
		{
			List<string> addedPlayers = new List<string>();
			foreach (var player in missingPlayers)
			{
				if (!addedPlayers.Contains(player.Value.Item1))
				{
					addedPlayers.Add(player.Value.Item1);

					var values = new Dictionary<string, string>
						{
							{ "gender", player.Value.Item2 }
						};

					var parts = player.Key.Split(",");
					string firstName = FormatName(parts[1]);
					string lastName = FormatName(parts[0]);

					var content = new StringContent(JsonSerializer.Serialize(values), Encoding.UTF8, "application/json");

					string url = $"https://tkhmiv70u9.execute-api.us-west-2.amazonaws.com/development/addPlayer/{firstName}/lastName/{lastName}";
					var response = client.PostAsync(url, content);
					response.Wait();
					string responseString = response.Result.Content.ReadAsStringAsync().Result;
					Console.WriteLine(responseString);
				}
			}
		}

		static void LoadNameData()
		{
			using (StreamReader reader = new StreamReader(@"C:\GitHub\FpaExcelResultsParser\playerNameData.json"))
			{
				playerNameData = JsonSerializer.Deserialize(reader.ReadToEnd(), typeof(PlayerDataCollection)) as PlayerDataCollection;
			}

			if (client == null)
			{
				HttpClientHandler handler = new HttpClientHandler()
				{
					AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
				};
				client = new HttpClient(handler);
			}
			HttpResponseMessage response = client.GetAsync("https://tkhmiv70u9.execute-api.us-west-2.amazonaws.com/development/getAllPlayers").Result;
			response.EnsureSuccessStatusCode();
			string result = response.Content.ReadAsStringAsync().Result;

			var responseData = JsonSerializer.Deserialize(result, typeof(PlayerDataResponse)) as PlayerDataResponse;
			internetPlayerData = responseData.players;

			foreach (var player in playerNameData)
			{
				bool found = false;
				foreach (var internetPlayer in internetPlayerData)
				{
					if (internetPlayer.Value.FullName == player.Value.Item1)
					{
						found = true;

						rawNameToId.Add(player.Key, internetPlayer.Key);
						break;
					}
				}

				if (!found)
				{
					missingPlayers.Add(player.Key, player.Value);
				}
			}
		}

		static void PrintNameData()
		{
			rawPlayerList.Sort();
			List<string> formattedNames = new List<string>();
			PlayerDataCollection nameDatabase = new PlayerDataCollection();
			foreach (var name in rawPlayerList)
			{
				var parts = name.Item1.Split(", ");
				if (parts.Length != 2 || parts[0].Trim().Length == 0 || parts[1].Trim().Length == 0)
				{
					Console.WriteLine(name);
				}

				string fullName = FormatFullName(name.Item1);
				if (!formattedNames.Contains(fullName))
				{
					formattedNames.Add(fullName);
				}

				nameDatabase.Add(name.Item1, new Tuple<string, string>(fullName, name.Item2?.ToUpper() ?? ""));
			}

			Console.WriteLine(JsonSerializer.Serialize(nameDatabase));
		}

		static string FormatFullName(string fullName)
		{
			var parts = fullName.Split(",");

			return $"{FormatName(parts[1])} {FormatName(parts[0])}";
		}

		static string FormatName(string name)
		{
			var lower = new StringBuilder(name.Trim().ToLower());
			for (int i = 1; i < lower.Length; ++i)
			{
				if (!char.IsLetter(lower[i - 1]))
				{
					lower[i] = char.ToUpper(lower[i]);
				}
			}

			lower[0] = char.ToUpper(lower[0]);

			return lower.ToString();
		}

		static int SearchForPlayers(ExcelWorksheet sheet, int column)
		{
			int lastRow = sheet.Dimension.End.Row;
			for (int row = 0; row <= lastRow; ++row)
			{
				string text = sheet.GetValue(row, column) as string;
				if (text == "Player")
				{
					return row;
				}
			}

			return -1;
		}

		static void ParsePlayers(ExcelWorksheet sheet, string filename)
		{
			int playerColumn = -1;
			int firstPlayerTagRow = -1;
			for (int column = 3; column < 10; ++column)
			{
				int row = SearchForPlayers(sheet, column);
				if (row >= 0)
				{
					firstPlayerTagRow = row;
					playerColumn = column;
					break;
				}
			}

			if (playerColumn >= 0)
			{
				int lastRow = sheet.Dimension.End.Row;
				for (int row = firstPlayerTagRow + 1; row <= lastRow; ++row)
				{
					string text = sheet.GetValue(row, playerColumn) as string;
					if (text != null && text.Length > 0 && text != "Player")
					{
						//if (text.Contains("Wegner (Enhuber?)"))
						//{
						//	Console.WriteLine(text + "  " + filename);
						//}

						if (rawPlayerList.FirstOrDefault(x => x.Item1 == text) == null && text.Contains(","))
						{
							rawPlayerList.Add(new Tuple<string, string>(text, sheet.GetValue(row, playerColumn + 1) as string));
						}
					}
				}
			}
			else
			{
				++errorCount;

				Console.WriteLine(filename);
			}
		}

		static ExcelWorksheet FindSheet(ExcelPackage package)
		{
			foreach (var sheet in package.Workbook.Worksheets)
			{
				if (sheet.Cells["B2"].Text.Contains("Event:"))
				{
					return sheet;
				}
			}

			return null;
		}

		static EventSummary ParseEventSummary(ExcelWorksheet sheet, string filename)
		{
			EventSummary summary = new EventSummary();
			summary.key = Guid.NewGuid().ToString();
			summary.eventName = sheet.Cells["I2"].Text;
			if (summary.eventName.Length == 0)
			{
				summary.eventName = sheet.Cells["D2"].Text;
			}
			if (summary.eventName.Length == 0)
			{
				summary.eventName = sheet.Cells["G2"].Text;
			}
			if (summary.eventName.Length == 0)
			{
				summary.eventName = sheet.Cells["J2"].Text;
			}
			if (summary.eventName.Length == 0)
			{
				summary.eventName = sheet.Cells["H2"].Text;
			}

			string dateString = sheet.Cells["I4"].Text;
			if (dateString.Length == 0)
			{
				dateString = sheet.Cells["D3"].Text;
			}
			if (dateString.Length == 0)
			{
				dateString = sheet.Cells["G4"].Text;
			}
			if (dateString.Length == 0)
			{
				dateString = sheet.Cells["J4"].Text;
			}
			if (dateString.Length == 0)
			{
				dateString = sheet.Cells["H4"].Text;
			}
			dateString = dateString.Replace(" ", "");
			var datePieces = dateString.Split('-');
			if (datePieces.Length == 2)
			{
				var startDatePieces = datePieces[0].Trim('.').Split('.');
				var endDatePiecies = datePieces[1].Trim('.').Split('.');
				if (startDatePieces.Length == 3 && endDatePiecies.Length == 3)
				{
					AssignDates(startDatePieces[0], startDatePieces[1], startDatePieces[2], endDatePiecies[0], endDatePiecies[1], endDatePiecies[2], filename, ref summary);
				}
				else if (startDatePieces.Length == 3 && endDatePiecies.Length == 2)
				{
					AssignDates(startDatePieces[0], startDatePieces[1], startDatePieces[2], endDatePiecies[0], endDatePiecies[1], startDatePieces[2], filename, ref summary);
				}
				else if (startDatePieces.Length == 2 && endDatePiecies.Length == 3)
				{
					AssignDates(startDatePieces[0], startDatePieces[1], endDatePiecies[2], endDatePiecies[0], endDatePiecies[1], endDatePiecies[2], filename, ref summary);
				}
				else if (startDatePieces.Length == 3 && endDatePiecies.Length == 1)
				{
					AssignDates(startDatePieces[0], startDatePieces[1], startDatePieces[2], startDatePieces[0], startDatePieces[1], endDatePiecies[0], filename, ref summary);
				}
				else if (startDatePieces.Length == 1 && endDatePiecies.Length == 3)
				{
					AssignDates(startDatePieces[0], endDatePiecies[1], endDatePiecies[2], endDatePiecies[0], endDatePiecies[1], endDatePiecies[2], filename, ref summary);
				}
				else
				{
					Console.WriteLine($"Error: Can't parse date string '{dateString}' from '{filename}'");
				}
			}
			else
			{
				var startDatePieces = dateString.Split('/');
				if (startDatePieces.Length == 3)
				{
					summary.startDate = $"{startDatePieces[2]}-{startDatePieces[0]}-{startDatePieces[1]}";
					summary.endDate = summary.startDate;
				}
				else
				{
					DateTime dateTime;
					if (DateTime.TryParse(dateString, out dateTime))
					{
						summary.startDate = $"{dateTime.Year}-{dateTime.Month}-{dateTime.Day}";
						summary.endDate = summary.startDate;
					}
					else
					{
						Console.WriteLine($"Error: Can't parse date string '{dateString}' from '{filename}'");
					}
				}
			}

			//var filePieces = filename.Replace(@"C:\GitHub\FpaExcelResultsParser\results\", "").Split(" - ");
			//string fileDate = "20" + filePieces[0].Substring(0, 2) + "-";
			//string numStr = filePieces[0].Substring(2, 2);
			//numStr = numStr[0] == '0' ? numStr.Substring(1, 1) : numStr;
			//fileDate += numStr + "-";
			//numStr = filePieces[0].Substring(4, 2);
			//numStr = numStr[0] == '0' ? numStr.Substring(1, 1) : numStr;
			//fileDate += numStr;
			//if (summary.startDate != fileDate)
			//{
			//	Console.WriteLine(summary.startDate + "  " + fileDate + "  " + filename);
			//}

			return summary;
		}

		static void LoadEvents()
		{
			if (client == null)
			{
				HttpClientHandler handler = new HttpClientHandler()
				{
					AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
				};
				client = new HttpClient(handler);
			}
			HttpResponseMessage response = client.GetAsync("https://xyf6qhiwi1.execute-api.us-west-2.amazonaws.com/development/getAllEvents").Result;
			response.EnsureSuccessStatusCode();
			string result = response.Content.ReadAsStringAsync().Result;

			var responseData = JsonSerializer.Deserialize(result, typeof(EventSummaryResponse)) as EventSummaryResponse;
			internetEventSummaries = responseData.allEventSummaryData;
		}

		static void AddEvents(List<EventSummary> summeries)
		{
			foreach (var summary in summeries)
			{
				if (!internetEventSummaries.Any(x => x.Value.Equals(summary)))
				{
					var values = new Dictionary<string, string>
						{
							{ "eventName", summary.eventName },
							{ "startDate", summary.startDate },
							{ "endDate", summary.endDate }
						};

					var content = new StringContent(JsonSerializer.Serialize(values), Encoding.UTF8, "application/json");

					string url = $"https://xyf6qhiwi1.execute-api.us-west-2.amazonaws.com/development/setEventSummary/{summary.key}";
					var response = client.PostAsync(url, content);
					response.Wait();
					string responseString = response.Result.Content.ReadAsStringAsync().Result;
					Console.WriteLine(responseString);
				}
			}
		}

		static void AssignDates(string start1Str, string start2Str, string start3Str, string end1Str, string end2Str, string end3Str, string filename, ref EventSummary summary)
		{
			int start1 = int.Parse(start1Str);
			int start2 = int.Parse(start2Str);
			int start3 = int.Parse(start3Str);
			int end1 = int.Parse(end1Str);
			int end2 = int.Parse(end2Str);
			int end3 = int.Parse(end3Str);

			TimeSpan shortestTimeSpan = new TimeSpan(100, 0, 0, 0);
			TimeSpan timeLength;
			DateTime startDateTime = new DateTime();
			DateTime endDateTime;
			int errorCount = 0;

			try
			{
				startDateTime = new DateTime(start1 < 100 ? start1 + 2000 : start1, start2, start3);
				endDateTime = new DateTime(end1 < 100 ? end1 + 2000 : end1, end2, end3);
				timeLength = endDateTime - startDateTime;
				if (timeLength < shortestTimeSpan)
				{
					shortestTimeSpan = timeLength;
					summary.startDate = $"{startDateTime.Year}-{startDateTime.Month}-{startDateTime.Day}";
					summary.endDate = $"{endDateTime.Year}-{endDateTime.Month}-{endDateTime.Day}";
				}
			}
			catch
			{
				++errorCount;
			}

			try
			{
				startDateTime = new DateTime(start3 < 100 ? start3 + 2000 : start3, start2, start1);
				endDateTime = new DateTime(end3 < 100 ? end3 + 2000 : end3, end2, end1);
				timeLength = endDateTime - startDateTime;
				if (timeLength < shortestTimeSpan)
				{
					shortestTimeSpan = timeLength;
					summary.startDate = $"{startDateTime.Year}-{startDateTime.Month}-{startDateTime.Day}";
					summary.endDate = $"{endDateTime.Year}-{endDateTime.Month}-{endDateTime.Day}";
				}
			}
			catch
			{
				++errorCount;
			}

			try
			{
				startDateTime = new DateTime(start1 < 100 ? start1 + 2000 : start1, start2, start3);
				endDateTime = new DateTime(end3 < 100 ? end3 + 2000 : end3, end2, end1);
				timeLength = endDateTime - startDateTime;
				if (timeLength < shortestTimeSpan)
				{
					shortestTimeSpan = timeLength;
					summary.startDate = $"{startDateTime.Year}-{startDateTime.Month}-{startDateTime.Day}";
					summary.endDate = $"{endDateTime.Year}-{endDateTime.Month}-{endDateTime.Day}";
				}
			}
			catch
			{
				++errorCount;
			}

			try
			{
				startDateTime = new DateTime(start3 < 100 ? start3 + 2000 : start3, start2, start1);
				endDateTime = new DateTime(end1 < 100 ? end1 + 2000 : end1, end2, end3);
				timeLength = endDateTime - startDateTime;
				if (timeLength < shortestTimeSpan)
				{
					shortestTimeSpan = timeLength;
					summary.startDate = $"{startDateTime.Year}-{startDateTime.Month}-{startDateTime.Day}";
					summary.endDate = $"{endDateTime.Year}-{endDateTime.Month}-{endDateTime.Day}";
				}
			}
			catch
			{
				++errorCount;
			}

			if (shortestTimeSpan > new TimeSpan(7, 0, 0, 0))
			{
				Console.WriteLine($"Error: Event too long. '{shortestTimeSpan.Days}' in '{filename}'");
			}
			else if (startDateTime > DateTime.Now)
			{
				Console.WriteLine($"Error: Event starting in the future. '{startDateTime.ToString()}' in '{filename}'");
			}
			else if (errorCount > 3)
			{
				Console.WriteLine($"Error: Can't parse date for '{filename}'");
			}
		}
	}
}
