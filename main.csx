#r "nuget: System.Management, 6.0.0"
using System.Management;

EnumUsbDevice(out var enumResult);

foreach(var drive in enumResult)
{
    var partitionNames = string.Join(",", drive.PartitionNameList);
    var driveLetters = string.Join(",", drive.DriveLetterList);
    Trace(TraceEventType.Information, $"{string.Join(", ", drive.Index, drive.PNPDeviceId, partitionNames, driveLetters)}");
}



void Trace(TraceEventType type, string msg){
    Console.WriteLine($"{type}, {msg}");
}

class DriveAndPartition
{
    public uint Index { get; set; }
    public string PNPDeviceId { get; set; }

    public List<string> PartitionNameList { get; set; } = new ();
    public List<string> DriveLetterList { get; set; } = new ();

}

private void EnumUsbDevice(out IEnumerable<DriveAndPartition> enumResult)
{
    var targetDiskIndexList = new List<DriveAndPartition>();

    {
        using var driveClass = new ManagementClass("Win32_DiskDrive");
        using var driveInstances = driveClass.GetInstances();
        foreach (var drive in driveInstances)
        {

            Trace(TraceEventType.Verbose, $"Win32_DiskDrive.GetInstances() Result, {drive}");

            if (!(drive.GetPropertyValue("InterfaceType") is string interfaceType))
            {
                continue;
            }

            if (interfaceType != "USB")
            {
                Trace(TraceEventType.Verbose, $"interfaceType!=USB, Skip, {interfaceType}");
                continue;
            }

            if (!(drive.GetPropertyValue("Index") is uint index))
            {
                Trace(TraceEventType.Verbose, $"Index Not Found, Skip, {drive}");
                continue;
            }

            if (!(drive.GetPropertyValue("Size") is ulong size))
            {
                Trace(TraceEventType.Verbose, $"Size Not Found, Skip, {drive}");
                continue;

            }

            if (!(drive.GetPropertyValue("PNPDeviceID") is string pnpDeviceId))
            {
                Trace(TraceEventType.Verbose, $"PNPDeviceID Not Found, Skip, {drive}");
                continue;
            }

            targetDiskIndexList.Add(new DriveAndPartition
            {
                Index = index,
                PNPDeviceId = pnpDeviceId,
            });

        }
    }

    {
        using ManagementClass partclass = new ManagementClass("Win32_DiskPartition");
        using ManagementObjectCollection partinstances = partclass.GetInstances();
        foreach (var part in partinstances)
        {
            using (part)
            {
                if (!(part.GetPropertyValue("DiskIndex") is uint diskIndex))
                {
                    Trace(TraceEventType.Verbose, $"DiskIndex Not Found, Skip, {part}");
                    continue;
                }
                if (!(part.GetPropertyValue("DeviceID") is string deviceId))
                {
                    Trace(TraceEventType.Verbose, $"DeviceID Not Found, Skip, {part}");
                    continue;
                }

                var target = targetDiskIndexList.FirstOrDefault(d => d.Index == diskIndex);
                if (target == null) continue;

                target.PartitionNameList.Add(part.ToString());

            }
        }
    }

    var driveLetterDict = new Dictionary<string, string>();

    using (ManagementClass ldclass = new ManagementClass("Win32_LogicalDisk"))
    using (ManagementObjectCollection ldinstances = ldclass.GetInstances())
    {
        foreach (var ld in ldinstances)
        {
            using (ld)
            {
                var driveLetter = ld.GetPropertyValue("DeviceID") as string;
                if (driveLetter != null)
                {
                    driveLetterDict.Add(ld.ToString(), driveLetter);
                }
            }
        }
    }

    using (ManagementClass vol2partclass = new ManagementClass("Win32_LogicalDiskToPartition"))
    using (ManagementObjectCollection vol2partinstances = vol2partclass.GetInstances())
    {
        foreach (var v2p in vol2partinstances)
        {
            using (v2p)
            {
                // 一致するロジカルディスクを取得
                var partComp = v2p.GetPropertyValue("Antecedent")?.ToString();
                var ldComp = v2p.GetPropertyValue("Dependent")?.ToString();
                if (partComp == null || ldComp == null) continue;

                var target = targetDiskIndexList.FirstOrDefault(d => d.PartitionNameList.Contains(partComp));
                if(target == null) continue;
                if (!driveLetterDict.TryGetValue(ldComp, out var driveValue)) continue;

                target.DriveLetterList.Add(driveValue);

            }
        }
    }

    enumResult = targetDiskIndexList;

}