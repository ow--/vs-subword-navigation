// Guids.cs
// MUST match guids.h
using System;

namespace VisualStudio.SubwordNavigation {
    static class GuidList {
        public const string guidSubwordNavigationPkgString = "ad7396e5-bd21-4114-90e6-36d157cbcc84";
        public const string guidSubwordNavigationCmdSetString = "693f5022-3bf5-419d-8bda-46f5d92b35ad";

        public static readonly Guid guidSubwordNavigationCmdSet = new Guid(guidSubwordNavigationCmdSetString);
    };
}