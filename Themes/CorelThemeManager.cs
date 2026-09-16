using System.Windows;

namespace SummaBarcodeCreater
{
    public class StylesController
    {
        string currentTheme = "";
        ResourceDictionary Resources;
        Corel.Interop.VGCore.Application corelApp;
        public StylesController(ResourceDictionary resource, Corel.Interop.VGCore.Application app)
        {
            Resources = resource;
            corelApp = app;
            
        }
        private readonly string[] StyleKeys = new string[] {
         "TabControl.Static.Border",
         "TabItem.Static.Border" ,
         "TabItem.Disabled.Background",
         "TabItem.Selected.Background",
         "TabItem.Static.Background",
         "TabItem.Selected.MouseOver.Background" ,
         "TabItem.Static.MouseOver.Background",
         "Button.MouseOver.Background" ,
         "Button.MouseOver.Border",
         "Button.Static.Border" ,
         "Button.Static.Background" ,
         "Button.Pressed.Background" ,
         "Button.Pressed.Border" ,
         "Button.Disabled.Foreground",
         "Button.Disabled.Background",
         "Default.Static.Foreground" ,
         "Default.Static.Background",
         "Container.Text.Static.Background" ,
         "Container.Text.Static.Foreground" ,
         "Container.Static.Background" ,
         "Default.Static.Inverted.Foreground" ,
         "ComboBox.Border.Popup.Item.MouseOver"
        };

        private void LoadStyle(string name)
        {

            string style = name.Substring(name.LastIndexOf("_") + 1);
            for (int i = 0; i < StyleKeys.Length; i++)
            {
                this.Resources[StyleKeys[i]] = this.Resources[string.Format("{0}.{1}", style, StyleKeys[i])];
            }
        }
      
        public void LoadThemeFromPreference()
        {
            try
            {
                string result = string.Empty;

                result = corelApp.GetApplicationPreferenceValue("WindowScheme", "Colors").ToString();


                if (!result.Equals(currentTheme))
                {
                    if (!result.Equals(string.Empty))
                    {
                        currentTheme = result;
                        LoadStyle(currentTheme);
                    }
                }
            }
            catch
            {
            }
        }

    }
}
