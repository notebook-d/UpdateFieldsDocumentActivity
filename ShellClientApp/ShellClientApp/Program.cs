using Ascon.Pilot.VirtualFileSystem.ShellServer;
using ShellClientApp;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0) return;    
            var path = args[0];                
            
            var address = ShellClientConstants.GetPilotPipeServerAddress(path, ShellClientConstants.COMMAND_PIPE_NAME);
            var command = new CommandInvokeData
            {
                CommandId = CommandIds.UpdateFilesAttributes,
                Paths = { ShellClientConstants.GetVirtualPath(path) }
            };
            var result = ShellClient.Request<CommandInvokeData, CommandInvokeResult>(command, address);

            Console.WriteLine(result.Result);
        }
    }
}