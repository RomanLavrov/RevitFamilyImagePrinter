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
				File.Create(_logFile);
			BeginLogSession();
		}

		private void BeginLogSession()
		{
			string paddingSymb = new string('~', 30);
			string header = $"{endl}{paddingSymb} LOGGER STARTED [{DateTime.Now.ToString()}] {paddingSymb}";
			File.AppendAllText(_logFile, header);
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

		public void Write(string text)
		{
			ConvertTabSymbols(ref text);
			text = text.Insert(0, $"[{DateTime.Now.ToString()}]: ");
			File.AppendAllText(_logFile, text);
		}

		public void WriteLine(string text)
		{
			ConvertTabSymbols(ref text);
			text = text.Insert(0, $"{endl}[{DateTime.Now.ToString()}]: ");
			File.AppendAllText(_logFile, $"{text}{endl}");
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
