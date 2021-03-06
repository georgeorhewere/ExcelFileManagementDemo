﻿
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.Caching.Memory;

namespace ExcelManagerLibrary
{
    public sealed class MemoryCacheManager
    {
        public static IMemoryCache MemoryCache { get; } = new MemoryCache(new MemoryCacheOptions());

    }
}
