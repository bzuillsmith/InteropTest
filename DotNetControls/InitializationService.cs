using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DotNetControls
{
    [ComVisible(true),
        Guid("6F970DF2-030C-45FC-802E-D6203D55693E")]
    public interface IInitializationService
    {
        void Initialize();
    }

    [ComVisible(true),
        Guid("F1471538-950D-489B-A610-5120931CC8F6"),
        ClassInterface(ClassInterfaceType.None)]
    public class InitializationService : IInitializationService
    {
        public void Initialize()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(true);
            }
            catch
            {
                // ignore
            }
        }
    }
}
