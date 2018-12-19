using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitFamilyImagePrinter.Infrastructure
{
	public class Logger
	{
		private readonly string _logFile;
		private static Logger _instance;
		private static readonly object _locker = new object();

		private Logger()
		{
			_logFile = Path.Combine(
				App.DefaultFolder, "log.txt");
			if (!File.Exists(_logFile))
				File.Create(_logFile);
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
			text = text.Insert(0, $"[{DateTime.Now.ToString()}]: ");
			File.AppendAllText(_logFile, text);
		}

		public void WriteLine(string text)
		{
			text = text.Insert(0, $"[{DateTime.Now.ToString()}]: ");
			File.AppendAllText(_logFile, $"{text}{Environment.NewLine}");
		}

		public void EndLogSession()
		{
			string endl = $"{Environment.NewLine}";
			string footer = $"{endl}{new string('=', 150)}{endl}{endl}{endl}";
			File.AppendAllText(_logFile, $"{footer}");
		}

		public void NewLine()
		{
			File.AppendAllText(_logFile, $"{Environment.NewLine}");
		}
	}
}
