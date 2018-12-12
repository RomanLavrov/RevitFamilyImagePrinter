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
		private readonly string logFile;
		private static Logger _instance;
		private static object _locker = new object();

		private Logger()
		{
			logFile = Path.Combine(Environment.CurrentDirectory, "log.txt");
			if (!File.Exists(logFile))
				File.Create(logFile);
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
			File.AppendAllText(logFile, text);
		}

		public void WriteLine(string text)
		{
			text = text.Insert(0, $"[{DateTime.Now.ToString()}]: ");
			File.AppendAllText(logFile, $"{text}{Environment.NewLine}");
		}

		public void EndLogSession()
		{
			string endl = $"{Environment.NewLine}";
			string footer = $"{endl}{new string('=', 150)}{endl}{endl}{endl}";
			File.AppendAllText(logFile, $"{footer}");
		}

		public void NewLine()
		{
			File.AppendAllText(logFile, $"{Environment.NewLine}");
		}
	}
}
