using System.Runtime.InteropServices;

namespace AresLicenseValidator.Interfaces
{
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789ABC")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAresLicenseValidator
    {
        [DispId(1)]
        bool ValidateLicense();

        [DispId(2)]
        string GetLicenseInfo();

        [DispId(3)]
        string GetLastError();

        [DispId(4)]
        string GetCurrentUser();

        [DispId(5)]
        int GetAuthorizedUserCount();
    }
}