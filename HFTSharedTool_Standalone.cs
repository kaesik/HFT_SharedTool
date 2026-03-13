namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            string Path = @"Z:\000_PMJ\Tekla\HFT_SharedTool\SharedTool\2024.0\HFT_SharedTool.exe";

            if (System.IO.File.Exists(exePath))
            {
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = exePath;

                process.StartInfo.Arguments = "standalone";

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