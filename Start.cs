using SummaBarcodeCreater;

namespace SummaBarcodeCreaterPlugin
{
	public partial class Main
	{
		private void Startup()
		{
		}

		[CgsAddInMacro]
		public void Start()
		{
			var entry = new Entry();
			entry.Begin(app);
		}

        [CgsAddInMacro]
        public void Settings()
		{
			var settings = new Settings_win();
			settings.Start(app);
		}
	}
}
