using XcelTools.Xtractor.Enums;

using System;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;

namespace XcelTools.Xtractor.Services
{
    //public static class ServiceProvider
    //{
    //    public static string Publisher => "XcelTools";

    //    public static string ApplicationName => "Xtractor";

    //    public static string AppDataDirectory => Path.Combine($@"C:\Users", Environment.UserName, @"AppData\Local", Publisher, ApplicationName);

    //    public static string ProgramsDirectory => @"C:\Xtractor";

    //    public static string DatabasesDirectory => Path.Combine(AppDataDirectory, "Databases");

    //    public static string JsonSectionsDirectory => Path.Combine(AppDataDirectory, @"Sections\Json");

    //    public static string JsonMaterialsDirectory => Path.Combine(AppDataDirectory, @"Materails\Json");

    //    public static string ThisPCMacAddress => NetworkInterface.GetAllNetworkInterfaces()
    //                                                             .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
    //                                                             .Select(nic => nic.GetPhysicalAddress().ToString())
    //                                                             .FirstOrDefault();

    //    public static string GetDataFileFullPath(Shape sectionType, FileExtension fileExtension = FileExtension.Json)
    //    {
    //        return Path.Combine(DatabasesDirectory, $"{sectionType}.{fileExtension.ToString().ToLower()}");
    //    }

    //    public static string LicenseFileFullPath => Path.Combine(AppDataDirectory, $"license.{FileExtension.Xtrlic.ToString().ToLower()}");
    //}

}
