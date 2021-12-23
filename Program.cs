using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace FpaExcelResultsParser
{
	class EventSummary
	{
		public string id;
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
			return (id + eventName + startDate + endDate).GetHashCode();
		}
	}

	class Program
	{
		static void Main(string[] args)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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
							var summaryData = ParseEventSummary(sheet, filename);
							//Console.WriteLine(JsonSerializer.Serialize(summaryData));

							if (!summeries.Contains(summaryData))
							{
								summeries.Add(summaryData);
							}

							if (summaryData.eventName.Length == 0)
							{
								Console.WriteLine(filename);
							}
						}
						else
						{
							Console.WriteLine($"Error: Can't find results sheet in '{filename}'");
						}
					}
				}
			}

			Console.WriteLine("Total events: " + summeries.Count);

			//using (var package = new ExcelPackage(new FileInfo(@"C:\GitHub\FpaExcelResultsParser\results\130105-V01-1302-In_den_Hallen.xlsx")))
			//{
			//	ExcelWorksheet sheet = package.Workbook.Worksheets[1];
			//	var summaryData = ParseEventSummary(sheet);
			//	Console.WriteLine(JsonSerializer.Serialize(summaryData));
			//}
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
			summary.id = Guid.NewGuid().ToString();
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
					AssignDates(endDatePiecies[0], endDatePiecies[1], startDatePieces[0], endDatePiecies[0], endDatePiecies[1], endDatePiecies[2], filename, ref summary);
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

			return summary;
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
			DateTime startDateTime;
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
			else if (errorCount > 3)
			{
				Console.WriteLine($"Error: Can't parse date for '{filename}'");
			}
		}
	}
}
