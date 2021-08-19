﻿using Microsoft.VisualStudio.Shell;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace VisualStudio.SubwordNavigation {
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "2.0.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidSubwordNavigationPkgString)]
    public sealed class SubwordNavigationPackage : AsyncPackage {
        public SubwordNavigationPackage() {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }
    }
}
