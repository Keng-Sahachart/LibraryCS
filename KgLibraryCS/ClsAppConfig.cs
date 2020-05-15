using Microsoft.VisualBasic;
using System;
using System.Configuration;

    // ##############################################################
    // ###    อย่าลืม Add reference System.Configuration            ####
    // ###    และ add new ไฟล์ app.config ด้วย                      ####
    // ###    และ ควรประกาศใช้ ClassAppConfig แค่ ตัวแปร เดียวเท่านั้น      ####
    // ##############################################################
    // วิธีใช้ใน VB
    // กรณี ใช้ กำหนดที่ Proerty
    // TextBox1.Text = My.Settings("Text_Logs") 

    // ในแท็ก ConnectionString
    // TextBox1.Text = ConfigurationManager.ConnectionStrings("Medline_DotNetConnectionString_OleDb").ConnectionString

    // ในแท็ก AppSetting
    // TextBox1.Text = ConfigurationManager.AppSettings("SQLconnection")
    // ####################################################################
namespace kgLibraryCs
{
    public class ClsAppConfig
    {
        private Configuration config =  ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        public AppSettingsSection AppSett;//= config.AppSettings;
        public ConnectionStringsSection ConnnectionStr ;//= config.ConnectionStrings;

        public ClsAppConfig()
        {
            ;
            //AppSett =  config.AppSettings;
            //ConnnectionStr = config.ConnectionStrings;
        }

        
        public void AppSetting_Set(string Key, string Value)
        {
             AppSettingsSection AppSett = config.AppSettings;
            // ถ้า Add รอบสอง จะเป็นการ ต่อสตริง
            AppSett.Settings.Add(Key, Value);
            // AppSett.Settings.Item("SomeKey").Value = 5 'just an example
            config.Save(ConfigurationSaveMode.Modified);
        }
        public string AppSetting_GetValue(string Key)
        {
            return AppSett.Settings[Key].Value;
        }
        public string AppSetting_GetValue(int Number)
        {
            return ConfigurationManager.AppSettings[Number];
        }

        public string AppSetting_GetKeyName(int Number)
        {
            return ConfigurationManager.AppSettings.Keys[Number];
        }
        public int AppSetting_Count()
        {
            return ConfigurationManager.AppSettings.Count;
        }

        public void AppSetting_Edit(string Key, string Value)
        {
            AppSett.Settings[Key].Value = Value;
            config.Save(ConfigurationSaveMode.Modified);
        }
        public void AppSetting_Edit2(string Key, string Value)
        {
            // AppSett.Settings.Add("Keyyy", "valueee")
            AppSett.Settings[Key].Value = Value; // just an example
                                                 // AppSett.Settings
            config.Save(ConfigurationSaveMode.Modified);
        }
        public bool AppSetting_KeyExists(string Key)
        {
            try
            {
                Information.IsNothing(AppSett.Settings[Key].Value);
                return true;
            }
            catch
            {
                return false;
            }
        }
        // ####################################################################
        // ####   Connection String                                       #####
        // ####################################################################
        public void ConnnectionString_Set(string Key, string Value)
        {
            try
            {
                var ConStrSett = new ConnectionStringSettings(Key, Value);
                ConnnectionStr.ConnectionStrings.Add(ConStrSett);
                config.Save(ConfigurationSaveMode.Modified);
            }
            catch //(Exception ex)
            {
                ConnnectionString_Edit(Key, Value);
            }
        }
        public string ConnnectionString_Get(string Key)
        {
            return ConnnectionStr.ConnectionStrings[Key].ToString();
        }
        public void ConnnectionString_Edit(string Key, string Value)
        {
            ConnnectionStr.ConnectionStrings[Key].ConnectionString = Value;
            config.Save(ConfigurationSaveMode.Modified);
        }
        public void ConnnectionString_Remove(string Key)
        {
            ConnnectionStr.ConnectionStrings.Remove(Key);
            config.Save(ConfigurationSaveMode.Modified);
        }
    }
}
