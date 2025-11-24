using System;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text;
using Ascon.Pilot.Core;
using ProtoBuf;

namespace Ascon.Pilot.VirtualFileSystem.ShellServer
{
    [ComVisible(false)]
    public static class ShellClient
    {
        private const int CONNECT_TIMEOUT = 10;

        public static TResponse Request<TRequest, TResponse>(TRequest request, string address) 
            where TRequest : class 
            where TResponse : class
        {
            return Request(request, 
                (s, r) => s.SerializeAndWrite(r),
                s => s.ReadAndDeserialize<TResponse>(), 
                address);
        }

        private static TResponse Request<TRequest, TResponse>(TRequest request, Action<PipeStream, TRequest> write, Func<PipeStream, TResponse> read, string address)
        {
            using (var client = new NamedPipeClientStream(".", address, PipeDirection.InOut))
            {
                try
                {
                    try
                    {
                        client.Connect(CONNECT_TIMEOUT);
                    }
                    catch (TimeoutException)
                    {
                        client.Connect(CONNECT_TIMEOUT);
                    }
                    catch (IOException)
                    {
                        client.Connect(CONNECT_TIMEOUT);
                    }
                }
                catch
                {
                    return default(TResponse);
                }

                write(client, request);
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    client.WaitForPipeDrain();
                return read(client);
            }
        }
    }

    [ComVisible(false)]
    public interface IPipeServerAddressProvider
    {
        string GetAddress(string path);
    }

    [ComVisible(false)]
    public class PipeServerAddressProvider : IPipeServerAddressProvider
    {
        public static IPipeServerAddressProvider IconAddressProvider = new PipeServerAddressProvider(ShellClientConstants.INFO_PIPE_NAME);
        public static IPipeServerAddressProvider CommandsAddressProvider = new PipeServerAddressProvider(ShellClientConstants.COMMAND_PIPE_NAME);

        private readonly string _pipeName;

        private PipeServerAddressProvider(string pipeName)
        {
            _pipeName = pipeName;
        }

        public string GetAddress(string path)
        {
            return ShellClientConstants.GetPilotPipeServerAddress(path, _pipeName);
        }
    }

    [ComVisible(false)]
    public static class PipeStreamExtensions
    {
        private const int INT_LENGTH = 4;
        private static readonly Encoding DefaultEncodingWindows = Encoding.Unicode;
        private static readonly Encoding DefaultEncodingLinux = Encoding.UTF8;

        public static void SerializeAndWrite<T>(this PipeStream stream, T instance) where T : class
        {
            using (var memoryStream = new MemoryStream())
            {
                Serializer.Serialize(memoryStream, instance);
                var buffer = memoryStream.ToArray();
                var lengthBytes = (byte[])new Int32Converter(buffer.Length);
                var result = new Byte[INT_LENGTH + buffer.Length];
                Buffer.BlockCopy(lengthBytes,0,result,0, lengthBytes.Length);
                Buffer.BlockCopy(buffer,0,result,lengthBytes.Length, buffer.Length);
                stream.Write(result, 0, result.Length);
                stream.Flush();
            }
        }

        public static T ReadAndDeserialize<T>(this PipeStream stream) where T : class
        {
            var lengthBytes = new byte[INT_LENGTH];
            if (stream.Read(lengthBytes, 0, INT_LENGTH) != INT_LENGTH)
                return null;

            var leng = new Int32Converter(lengthBytes);
            var buffer = new byte[leng];
            stream.Read(buffer, 0, buffer.Length);
            using (var memoryStream = new MemoryStream(buffer))
            {
                return Serializer.Deserialize<T>(memoryStream);
            }
        }

        public static string ReadString(this PipeStream stream)
        {
            var first = stream.ReadByteSafe();
            var second = stream.ReadByteSafe();
            int length = first * 256 + second;
            var inBuffer = new byte[length];
            stream.Read(inBuffer, 0, length);
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? DefaultEncodingWindows.GetString(inBuffer) : DefaultEncodingLinux.GetString(inBuffer, 0, length);
        }

        public static void WriteString(this PipeStream stream, string str)
        {
            byte[] buffer;
            buffer = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? DefaultEncodingWindows.GetBytes(str) : DefaultEncodingLinux.GetBytes(str);
            int len = Math.Min(buffer.Length, UInt16.MaxValue);

            var first = (byte)(len / 256);
            var second = (byte)(len & 255);
            
            var result = new byte[len + 2];
            result[0] = first;
            result[1] = second;
            Buffer.BlockCopy(buffer,0,result,2,len);
            stream.Write(result, 0, len+2);
            stream.Flush();
        }

        public static int ReadByteSafe(this PipeStream stream)
        {
            int result = -1;
            while (result == -1)
                result = stream.ReadByte();
            return result;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct Int32Converter
        {
            [FieldOffset(0)]
            private readonly int _value;
            [FieldOffset(0)]
            private readonly byte _byte1;
            [FieldOffset(1)]
            private readonly byte _byte2;
            [FieldOffset(2)]
            private readonly byte _byte3;
            [FieldOffset(3)]
            private readonly byte _byte4;

            public Int32Converter(int value)
            {
                _byte1 = _byte2 = _byte3 = _byte4 = 0;
                _value = value;
            }

            public Int32Converter(byte[] value)
            {
                _value = 0;
                _byte1 = value[0];
                _byte2 = value[1];
                _byte3 = value[2];
                _byte4 = value[3];
            }

            public static implicit operator int(Int32Converter value)
            {
                return value._value;
            }

            public static implicit operator byte[] (Int32Converter value)
            {
                return new[] { value._byte1, value._byte2, value._byte3, value._byte4 };
            }

            public static implicit operator Int32Converter(int value)
            {
                return new Int32Converter(value);
            }
        }
    }
}
