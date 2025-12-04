namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            string XS_Variable = System.Environment.GetEnvironmentVariable("XSDATADIR");
            string TS_Plugin = @"\environments\common\extensions\SharedTool\";
            string TS_Application = "HFT_SharedTool.exe";

            string exePath = XS_Variable + TS_Plugin + TS_Application;

            if (System.IO.File.Exists(exePath))
            {
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = exePath;

                process.StartInfo.Arguments = "writeout";

                process.Start();
                process.Close();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(
                    TS_Application + " not found!\n\nCheck: " + XS_Variable + TS_Plugin,
                    "Tekla Structures",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }
    }
}