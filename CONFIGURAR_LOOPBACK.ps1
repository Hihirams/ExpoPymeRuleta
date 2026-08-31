# ============================================================
#  Expo PyME - Configurar adaptador Loopback para Mobile Hotspot
#  (Opcion 2: crear red WiFi local SIN internet en la laptop)
#
#  Que hace:
#   1) Se auto-eleva a Administrador (aparece un aviso: dile "Si").
#   2) Instala el "Microsoft KM-TEST Loopback Adapter" (sin devcon).
#   3) Lo renombra a "HotspotLoopback" y le da una IP propia para que
#      el Mobile Hotspot lo pueda usar como origen a compartir.
#   4) Deja un registro en: configurar_loopback.log
#
#  Despues de correrlo: Configuracion > Mobile Hotspot >
#     "Compartir mi conexion desde" = HotspotLoopback  >  Encender.
# ============================================================

$ErrorActionPreference = "Stop"
$LogPath = Join-Path $PSScriptRoot "configurar_loopback.log"

function Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $line
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

# ---- 1) Auto-elevacion -------------------------------------------------------
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin = (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Solicitando permisos de administrador (di 'Si' en el aviso)..."
    Start-Process -FilePath "powershell.exe" `
        -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","`"$PSCommandPath`"" `
        -Verb RunAs
    exit
}

"" | Set-Content -Path $LogPath -Encoding UTF8   # reinicia el log
Log "=== Configuracion de Loopback iniciada (Administrador) ==="

# ---- 2) Ya existe? -----------------------------------------------------------
$existing = Get-NetAdapter -ErrorAction SilentlyContinue |
    Where-Object { $_.InterfaceDescription -like "*Loopback*" }

if ($existing) {
    Log "El adaptador Loopback ya existe: '$($existing.Name)'. Saltando instalacion."
} else {
    Log "Instalando el adaptador Loopback (Microsoft KM-TEST)..."

    $inf = Join-Path $env:windir "INF\netloop.inf"
    if (-not (Test-Path $inf)) { Log "ERROR: no se encontro $inf"; Read-Host "Enter para salir"; exit 1 }

    # Instala el devnode con hardware id *MSLOOP usando SetupAPI (lo mismo que
    # hace devcon internamente). No requiere descargas.
    $code = @'
using System;
using System.Runtime.InteropServices;
public static class LoopInstall {
    const uint DICD_GENERATE_ID = 0x1;
    const uint SPDRP_HARDWAREID = 0x1;
    const uint DIF_REGISTERDEVICE = 0x19;
    const uint INSTALLFLAG_FORCE = 0x1;
    [StructLayout(LayoutKind.Sequential)]
    struct SP_DEVINFO_DATA { public uint cbSize; public Guid ClassGuid; public uint DevInst; public IntPtr Reserved; }
    [DllImport("setupapi.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern IntPtr SetupDiCreateDeviceInfoList(ref Guid ClassGuid, IntPtr hwndParent);
    [DllImport("setupapi.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern bool SetupDiCreateDeviceInfoW(IntPtr DeviceInfoSet, string DeviceName, ref Guid ClassGuid, string DeviceDescription, IntPtr hwndParent, uint CreationFlags, ref SP_DEVINFO_DATA DeviceInfoData);
    [DllImport("setupapi.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern bool SetupDiSetDeviceRegistryPropertyW(IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData, uint Property, byte[] PropertyBuffer, uint PropertyBufferSize);
    [DllImport("setupapi.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern bool SetupDiCallClassInstaller(uint InstallFunction, IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData);
    [DllImport("setupapi.dll", SetLastError=true)]
    static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);
    [DllImport("newdev.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern bool UpdateDriverForPlugAndPlayDevicesW(IntPtr hwndParent, string HardwareId, string FullInfPath, uint InstallFlags, out bool bRebootRequired);

    public static int Install(string infPath, string hwid) {
        Guid netClass = new Guid("4d36e972-e325-11ce-bfc1-08002be10318");
        IntPtr set = SetupDiCreateDeviceInfoList(ref netClass, IntPtr.Zero);
        if (set == (IntPtr)(-1)) return Err("CreateDeviceInfoList");
        try {
            SP_DEVINFO_DATA dev = new SP_DEVINFO_DATA();
            dev.cbSize = (uint)Marshal.SizeOf(dev);
            if (!SetupDiCreateDeviceInfoW(set, "NET", ref netClass, null, IntPtr.Zero, DICD_GENERATE_ID, ref dev))
                return Err("CreateDeviceInfo");
            byte[] buf = System.Text.Encoding.Unicode.GetBytes(hwid + "\0\0");
            if (!SetupDiSetDeviceRegistryPropertyW(set, ref dev, SPDRP_HARDWAREID, buf, (uint)buf.Length))
                return Err("SetHardwareId");
            if (!SetupDiCallClassInstaller(DIF_REGISTERDEVICE, set, ref dev))
                return Err("RegisterDevice");
            bool reboot;
            if (!UpdateDriverForPlugAndPlayDevicesW(IntPtr.Zero, hwid, infPath, INSTALLFLAG_FORCE, out reboot))
                return Err("UpdateDriver");
            return 0;
        } finally { SetupDiDestroyDeviceInfoList(set); }
    }
    static int Err(string where) {
        int e = Marshal.GetLastWin32Error();
        Console.Error.WriteLine("Fallo en " + where + " (codigo " + e + ")");
        return e == 0 ? -1 : e;
    }
}
'@
    $ok = $false
    try {
        Add-Type -TypeDefinition $code -Language CSharp
        $rc = [LoopInstall]::Install($inf, "*MSLOOP")
        if ($rc -eq 0) { $ok = $true; Log "Instalacion headless OK." }
        else { Log "Instalacion headless devolvio codigo $rc." }
    } catch {
        Log "Instalacion headless no disponible: $($_.Exception.Message)"
    }

    if (-not $ok) {
        Log "Abriendo el asistente manual (Agregar hardware heredado)..."
        Log "  En el asistente: Instalar manualmente > Adaptadores de red >"
        Log "  Fabricante Microsoft > 'Microsoft KM-TEST Loopback Adapter' > Instalar."
        Log "  Al terminar, vuelve a ejecutar este archivo para configurarlo."
        Start-Process "hdwwiz.exe" -Verb RunAs
        Read-Host "Presiona Enter para cerrar"
        exit 2
    }
}

# ---- 3) Configurar el adaptador ---------------------------------------------
Start-Sleep -Seconds 2
$lb = Get-NetAdapter -ErrorAction SilentlyContinue |
    Where-Object { $_.InterfaceDescription -like "*Loopback*" } | Select-Object -First 1

if (-not $lb) { Log "ERROR: no encuentro el adaptador tras instalar."; Read-Host "Enter para salir"; exit 1 }

$targetName = "HotspotLoopback"
if ($lb.Name -ne $targetName) {
    try { Rename-NetAdapter -Name $lb.Name -NewName $targetName -ErrorAction Stop; Log "Renombrado a '$targetName'." }
    catch { Log "No se pudo renombrar (no es critico): $($_.Exception.Message)"; $targetName = $lb.Name }
}

# IP propia en una subred distinta a la del hotspot (192.168.137.x) para que
# el Mobile Hotspot lo acepte como origen sin conflictos.
try {
    Get-NetIPAddress -InterfaceAlias $targetName -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notlike "169.254.*" } |
        Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue
    New-NetIPAddress -InterfaceAlias $targetName -IPAddress "192.168.199.1" -PrefixLength 24 -ErrorAction Stop | Out-Null
    Log "IP asignada al loopback: 192.168.199.1/24"
} catch {
    Log "Aviso al asignar IP (no es critico): $($_.Exception.Message)"
}

Log ""
Log "======================================================"
Log " LISTO. Adaptador '$targetName' preparado."
Log "======================================================"
Log " Ahora abre:  Configuracion > Red e Internet > Mobile hotspot"
Log "   1) 'Compartir mi conexion de Internet desde' -> $targetName"
Log "   2) 'Compartir a traves de' -> Wi-Fi"
Log "   3) Enciende el interruptor (Off -> On)"
Log ""
Log " Luego inicia el servidor con INICIAR_SERVIDOR.bat y muestra el QR"
Log " del panel. La laptop sera 192.168.137.1 para los celulares."
Log "======================================================"

# Abre directamente la pagina de Mobile Hotspot para el ultimo paso.
try { Start-Process "ms-settings:network-mobilehotspot" } catch {}

Read-Host "Presiona Enter para cerrar"
