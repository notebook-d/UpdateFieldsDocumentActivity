using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Ascon.Pilot.VirtualFileSystem.ShellServer
{
    [ComVisible(false)]
    public static class ShellClientConstants
    {
        public const string INFO_PIPE_NAME = "PilotInfoPipe";
        public const string COMMAND_PIPE_NAME = "PilotCommandPipe";
        public const string FILES_SEPARATOR = ">";
        public const string VIRTUAL_DRIVE_FORMAT_NAME = "pfmfs.";
        public const string NTFS_DRIVE_FORMAT_NAME = "NTFS";
        public const string PROTOCOL_VERSION = "20170614";

        public static string GetPilotPipeServerAddress(string path, string pipeName)
        {
            try
            {
                if (string.IsNullOrEmpty(path))
                    return null;

                if (!char.IsLetter(path, 0))
                    return null;

                var root = Path.GetPathRoot(path);
                if (string.IsNullOrEmpty(root))
                    return null;

                var driveInfo = new DriveInfo(root);
                if (driveInfo.DriveType != DriveType.Fixed)
                    return null;

                var driveFormat = driveInfo.DriveFormat;
                return driveFormat.StartsWith(VIRTUAL_DRIVE_FORMAT_NAME)
                    ? pipeName + Environment.UserName + driveInfo.DriveFormat + PROTOCOL_VERSION
                    : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool IsVirtualDrivePath(string path)
        {
            try
            {
                return GetPilotPipeServerAddress(path, string.Empty) != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsVirtualDrivePathFor(this string path, string clientName)
        {
            try
            {
                var serverName = GetPilotPipeServerAddress(path, string.Empty);
                return serverName != null && serverName.EndsWith(clientName + PROTOCOL_VERSION);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static string GetVirtualPath(string path)
        {
            string virtualPath;
            if (Environment.OSVersion.Platform == PlatformID.Unix)
                virtualPath = Path.Combine(path.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries).ToList()
                    .SkipWhile(x => !x.Contains("3D")).Skip(1).ToArray());
            else
                virtualPath = path?.Substring(2) ?? string.Empty;

            if (virtualPath.StartsWith("\\\\"))
                virtualPath = virtualPath.Substring(2);
            if (virtualPath.StartsWith("\\"))
                virtualPath = virtualPath.Substring(1);
         
            return virtualPath;
        }

        public static string BuildServerAddress(string serverName, string clientName)
        {
            return $"{serverName}{Environment.UserName}{VIRTUAL_DRIVE_FORMAT_NAME}{clientName}{PROTOCOL_VERSION}";
        }

        public static string BuildServerAddressLinux(string serverName, string clientName)
        {
            return $"{serverName}{VIRTUAL_DRIVE_FORMAT_NAME}{clientName}{PROTOCOL_VERSION}";
        }
    }
}
