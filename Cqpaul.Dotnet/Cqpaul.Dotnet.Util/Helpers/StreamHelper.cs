namespace Cqpaul.Dotnet.Util.Helpers
{
    public static class StreamHelper
    {
        public static byte[] StreamToBytes(Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;
        }

        public static Stream BytesToStream(byte[] bytes)
        {
            Stream stream = new MemoryStream(bytes);
            return stream;
        }

        public static string StreamToString(Stream fileStream)
        {
            StreamReader streamReader = new StreamReader(fileStream);
            string text = streamReader.ReadToEnd();
            return text;
        }
    }
}
