using OpenMcdf;

namespace ExcelFileTools
{
    public class CompoundStream : Stream
    {
        private readonly CFStream _cfStream;

        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => _cfStream != null;

        public override long Length => throw new NotImplementedException();

        public override long Position { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public CompoundStream(CFStream cfStream)
        {
            _cfStream = cfStream;
        }

        public override void Flush()
        {

        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotImplementedException();
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _cfStream.Append(buffer.Skip(offset).Take(count).ToArray());
        }
    }
}
