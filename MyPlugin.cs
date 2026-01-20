
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace GM.XrmToolBox.UserRoleMatrix
{
    [Export(typeof(IXrmToolBoxPlugin))]
    [ExportMetadata("Name", "User Roles Matrix")]
    [ExportMetadata("Description", "Lists users and their security roles (direct + via Owner Teams), including Business Unit info, filters, duplicate detection, and export.")]
    [ExportMetadata("SmallImageBase64", "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAADf0lEQVR42u2Xa0hTYRjH/2eXc3brzNXUuUJDaxOnrVgXukmEZZkfihYliRRBYWQWRCn1qaB1nUVhsT5UUISaIFnR5UPQfUZmMaNsoEsj03ltc3q2efrQWrNpTmquDz1wvpyH9/3/zvM+z3OeF4iwESN6iizsX1U6kkqEBlBkYVmDJiEsX1tcb/sVhDecOFFcbwsHAGvQJBCwsIEQnPES/xEB1qBJCDxePwnLsvHhFA+KBEF8/BmBv51woZhPkxPpMvwPEHEA3kgOikcgf95ErEujoYmlICY5cDCDaHd48d4+gJxrLXAwg7izOQGZKgkAoMHOINn4AawvpUkugaZ9KsRN+C5zta4HuWUto0dALubCvD0RJdkK8Fwd3cv1m0olsuhNM1NU+cUFW05L7PUN1O19RjB9/YHrVHISWY3Gc/AhbNBK/eIAgPqbD1FdVDJqBExrlNDGCeDsczHZ6brCdp4iFmvO7rfFqKfaHPbuSuNzK9pau+Bl3AAEAMAwjIckSd6uLTkLbpke1SIpXbdr4SQE+kI6glgJD6tTaABARXnZ4/ZuB4P80gMQyb6/lCqjIVVGQ71sfuC6urq6Rqk0SpyRkaHVHL9YGpO4UjdLKYC59s3HKDHFVavVk0NKwpQYCoSvP1qt1s9IzlwAkYy+sn4KWIPG/+xZLB/a3ViWPX2p3AwAhfol03fPFbgAoOTY4QrwKDLkKiCIIZsCsvg4AMgta4FQt/aM3/ngxGWYL1UFrr1cUf2yy+Fy5+XlLVmVFito/tzWW1lZ+QwCWhIywNu2AX8WJyUlKYYQZR0q+F1J9bn63Rde9HAoiuJzOARx5tTJKo9ILgMpFoYM0PrVg2pLhxsA9Hr9/ChPR9dY6vqsuZfr8Q6yTqez/4LJdA+6javG3Ii23rDz31ptdpqmRVWGbYu0E70ukksgTSEYFaC5xw3+zgcdElqa2+0c8EKrXzbmRvTF4cGc8838Amn5jXVZSzWPt0/jCoRCtrd/kGho/eq2PL3/sqam5gO4c2cPuyutkGPv6+t/1An7+DLp0Z6lmUf3XL2Fd3fPobPpEzwDbpASIUQyGtHT4jE7dfqKizagbNtBND55BeUM1XB7JRutgCl7BzqbPkGTnR48E47DNBQ0kPjmw/9/w38EYIRLQ1jNp8kJGpnHKwGDjuBIKhFuiMDs/2euZhG/nEbcvgEhmWjf6ekl5gAAAABJRU5ErkJggg==")]
    [ExportMetadata("BigImageBase64", "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")]
    [ExportMetadata("BackgroundColor", "White")]
    [ExportMetadata("PrimaryFontColor", "Black")]
    [ExportMetadata("SecondaryFontColor", "Gray")]
    public class MyPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new MyPluginControl();
        }

        /// <summary>
        /// Constructor 
        /// </summary>
        public MyPlugin()
        {
            // If you have external assemblies that you need to load, uncomment the following to 
            // hook into the event that will fire when an Assembly fails to resolve
            // AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            // check to see if the failing assembly is one that we reference.
            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly != null)
            {
                // load from the path to this plugin assembly, not host executable
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }

    }
}
