// Program.cs

using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AmqpModbusIntegration
{
    internal class Program
    {
        static ModbusViewer modbusViewer = new ModbusViewer();

        [STAThread]
        static void Main()
        {
            Task.Run(() => AmqpEndpointHandler.StartAmqpEndpoint(modbusViewer));  // 將AMQP端點啟動在非同步任務中
            Application.Run(modbusViewer);        // 啟動UI主線程
        }
    }
}
