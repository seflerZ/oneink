namespace OneInk
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices.ComTypes;
    using Marshal = System.Runtime.InteropServices.Marshal;

    internal class ReadOnlyStream : IStream, IDisposable
    {
        private Stream stream;

        public ReadOnlyStream(Stream wrapper)
        {
            stream = wrapper;
        }

        #region IDisposable Support
        private bool disposed = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    if (stream != null)
                    {
                        stream.Dispose();
                        stream = null;
                    }
                }
                disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region IStream Members

        public void Clone(out IStream ppstm)
        {
            ppstm = new ReadOnlyStream(stream);
        }

        public void Commit(int grfCommitFlags)
        {
            stream.Flush();
        }

        public void CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten)
        {
            // N/A
        }

        public void LockRegion(long libOffset, long cb, int dwLockType)
        {
            throw new NotSupportedException();
        }

        public void Read(byte[] pv, int cb, IntPtr pcbRead)
        {
            Marshal.WriteInt64(pcbRead, (long)stream.Read(pv, 0, cb));
        }

        public void Revert()
        {
            throw new NotSupportedException();
        }

        public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
        {
            long num;
            switch (dwOrigin)
            {
                case 0: // STREAM_SEEK_SET
                    num = dlibMove;
                    break;
                case 1: // STREAM_SEEK_CUR
                    num = stream.Position + dlibMove;
                    break;
                case 2: // STREAM_SEEK_END
                    num = stream.Length + dlibMove;
                    break;
                default:
                    return;
            }

            if ((num >= 0L) && (num < stream.Length))
            {
                stream.Position = num;
            }
            Marshal.WriteInt64(plibNewPosition, stream.Position);
        }

        public void SetSize(long libNewSize)
        {
            stream.SetLength(libNewSize);
        }

        public void Stat(out STATSTG pstatstg, int grfStatFlag)
        {
            pstatstg = new STATSTG { cbSize = stream.Length };
        }

        public void UnlockRegion(long libOffset, long cb, int dwLockType)
        {
            throw new NotSupportedException();
        }

        public void Write(byte[] pv, int cb, IntPtr pcbWritten)
        {
            Marshal.WriteInt64(pcbWritten, 0L);
            stream.Write(pv, 0, cb);
            Marshal.WriteInt64(pcbWritten, (long)cb);
        }

        #endregion
    }
}
