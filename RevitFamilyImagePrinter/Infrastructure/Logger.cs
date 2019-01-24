using System;
using System.IO;

namespace RevitFamilyImagePrinter.Infrastructure
{
	public class Logger
	{
		private readonly string _logFile;
		private readonly string endl = $"{Environment.NewLine}";
		private static Logger _instance;
		private static readonly object _locker = new object();

		private Logger()
		{
			_logFile = Path.Combine(
				App.DefaultFolder, "log.txt");
			if (!File.Exists(_logFile))
			{
				using (File.Create(_logFile)) { }
			}
			BeginLogSession();
		}

		private void BeginLogSession()
		{
			string paddingSymb = new string('~', 30);
			string pcInfo =
				$"{endl}{endl}Operating system: {GetOsName()} ({(Environment.Is64BitOperatingSystem ? "x64" : "x86")})";
			string header = $"{endl}{paddingSymb} LOGGER STARTED [{DateTime.Now.ToString()}] {paddingSymb}{pcInfo}";
			File.AppendAllText(_logFile, header);
		}

		private string GetOsName()
		{
			OperatingSystem osInfo = System.Environment.OSVersion;
			string version = $"{osInfo.Version.Major}.{osInfo.Version.Minor}";
			switch (version)
			{
				case "10.0": return "Windows 10 / Server 2016";
				case "6.3": return "Windows 8.1 / Server 2012 R2";
				case "6.1": return "Windows 7 / Server 2008 R2";
				case "6.0": return "Server 2008 / Windows Vista";
				case "5.2": return "Server 2003 R2 / Server 2003 / XP 64-Bit Edition";
			}
			return "Unknown";
		}

		private void ConvertTabSymbols(ref string text)
		{
			text = text.Replace("\n", "\r\n");
			text = text.Replace("\t", "\r\t");
		}

		public static Logger GetLogger()
		{
			lock (_locker)
			{
				if (_instance == null)
					_instance = new Logger();
			}
			return _instance;
		}

		public void Write(string text, bool isTimeLog = true)
		{
			ConvertTabSymbols(ref text);
			if (isTimeLog) text = text.Insert(0, $"[{DateTime.Now.ToString()}]: ");
			File.AppendAllText(_logFile, text);
		}

		public void WriteLine(string text, bool isTimeLog = true)
		{
			ConvertTabSymbols(ref text);
			if (isTimeLog) text = text.Insert(0, $"[{DateTime.Now.ToString()}]: ");
			File.AppendAllText(_logFile, $"{endl}{text}{endl}");
		}

		public void EndLogSession()
		{
			string paddingSymb = new string('=', 30);
			string footer = $"{endl}{paddingSymb} LOGGER FINISHED [{DateTime.Now.ToString()}] {paddingSymb}";
			File.AppendAllText(_logFile, $"{footer}");
		}

		public void NewLine()
		{
			File.AppendAllText(_logFile, $"{Environment.NewLine}");
		}
	}
}
