using System.Text;

namespace ExcelFileTools
{
    public static class StringUtils
    {
        public static void WriteAsNullTerminatedAsciiWithLenPrefixString(this string value, BinaryWriter binaryWriter)
        {
            var fileNameBytes = Encoding.ASCII.GetBytes(value);

            binaryWriter.Write((UInt32)fileNameBytes.Length + 1);
            binaryWriter.Write(fileNameBytes, 0, fileNameBytes.Length);
            binaryWriter.Write((byte)0);
        }

        public static void WriteAsNullTerminatedAsciiString(this string value, BinaryWriter binaryWriter)
        {
            var fileNameBytes = Encoding.ASCII.GetBytes(value);
            binaryWriter.Write(fileNameBytes, 0, fileNameBytes.Length);
            binaryWriter.Write((byte)0);
        }

        public static void WriteAsSizeAsciiPrefixedString(this string value, BinaryWriter binaryWriter)
        {
            var fileNameBytes = Encoding.ASCII.GetBytes(value);
            binaryWriter.Write((UInt32)(fileNameBytes.Length + 1));
            binaryWriter.Write(fileNameBytes, 0, fileNameBytes.Length);
            binaryWriter.Write((byte)0);
        }

    }
}
