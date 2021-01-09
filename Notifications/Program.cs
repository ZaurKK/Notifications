using Microsoft.Extensions.CommandLineUtils;

namespace Notifications
{
    class Program
    {
        static DPO DPO { get; } = new DPO();

        static void Main(string[] args)
        {
            var cmd = new CommandLineApplication();
            var argScanFolder = cmd.Option("-scan_folder <value>", "Folder to scan for DPL archives and notification files", CommandOptionType.SingleValue);
            var argForce = cmd.Option("-force", "Force to recreate folders and archives", CommandOptionType.NoValue);
            //var argDPL = cmd.Option("-dpl <value>", "DPL archive file path", CommandOptionType.SingleValue);
            //var argNotify = cmd.Option("-notify <value>", "Path to excel file with noifications", CommandOptionType.SingleValue);

            var scanFolder = "";
            bool force = false;
            cmd.OnExecute(() =>
            {
                scanFolder = argScanFolder.Value();
                force = argForce.HasValue();
                //dplArchivePath = argDPL.Value();
                //notifyFilePath = argNotify.Value();
                return 0;
            });

            cmd.HelpOption("-? | -h | --help");
            cmd.Execute(args);

            DPO.Run(scanFolder, force);
        }
    }
}
