namespace ZipHelper.Zip.Compression
{
    public class DeflaterPending : PendingBuffer
    {
        /// <summary>
        /// Construct instance with default buffer size
        /// </summary>
        public DeflaterPending() : base(DeflaterConstants.PENDING_BUF_SIZE)
        {
        }
    }
}