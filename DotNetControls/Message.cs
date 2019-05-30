using System;
using System.Runtime.InteropServices;

namespace DotNetControls
{
    [ComVisible(true),
        Guid("C6DF0D2C-8062-40A0-9D8E-F061AFEA9EDA")]
    public interface IMyTestClass
    {
        string Message { get; set; }
    }

    [ComVisible(true),
        Guid("B7F1224D-D2AF-40D0-8F81-02E17686625E"),
        ClassInterface(ClassInterfaceType.None)]
    public class MyTestClass : IMyTestClass
    {
        public string Message { get; set; }
    }
}
