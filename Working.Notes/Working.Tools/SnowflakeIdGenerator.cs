using System;

namespace Working.Tools
{

    /// <summary>
    /// 雪花ID生成类
    /// </summary>
    public static class SnowflakeIdGenerator
    {
        private static readonly long Epoch = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks;
        private const int TimestampBits = 41;
        private const int DatacenterIdBits = 5;
        private const int WorkerIdBits = 5;
        private const int SequenceBits = 12;

        private static long _lastTimestamp = -1L;
        private static long _sequence = 0L;

        private static readonly object _lock = new object();

        /// <summary>
        /// 创建一个雪花ID
        /// </summary>
        /// <param name="datacenterId">数据中心ID</param>
        /// <param name="workerId">工作节点ID</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public static long GenerateId(int datacenterId, int workerId)
        {
            lock (_lock)
            {
                long timestamp = DateTime.UtcNow.Ticks - Epoch;

                if (timestamp < _lastTimestamp)
                {
                    throw new InvalidOperationException("Invalid system clock.");
                }

                if (timestamp == _lastTimestamp)
                {
                    _sequence = (_sequence + 1) & ((1 << SequenceBits) - 1);

                    if (_sequence == 0)
                    {
                        timestamp = WaitNextMillis(_lastTimestamp);
                    }
                }
                else
                {
                    _sequence = 0L;
                }

                _lastTimestamp = timestamp;

                long id = ((timestamp << (DatacenterIdBits + WorkerIdBits + SequenceBits)) |
                           ((long)datacenterId << (WorkerIdBits + SequenceBits)) |
                           ((long)workerId << SequenceBits) |
                           _sequence);

                return id;
            }
        }

        /// <summary>
        /// 创建一个雪花ID  数据中心ID=1 工作节点Id=1
        /// </summary>
        /// <returns></returns>
        public static long CreateNextId()
        {
           return GenerateId(1,1);
        }

        private static long WaitNextMillis(long lastTimestamp)
        {
            long timestamp = DateTime.UtcNow.Ticks - Epoch;
            while (timestamp <= lastTimestamp)
            {
                timestamp = DateTime.UtcNow.Ticks - Epoch;
            }
            return timestamp;
        }
    }

}
