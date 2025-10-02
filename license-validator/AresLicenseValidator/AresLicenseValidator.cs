using System;
using System.Runtime.InteropServices;
using AresLicenseValidator.Interfaces;
using AresLicenseValidator.Services;
using AresLicenseValidator.Models;

namespace AresLicenseValidator
{
    [ComVisible(true)]
    [Guid("87654321-4321-4321-4321-CBA987654321")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ARES.LicenseValidator")]
    public class AresLicenseValidator : IAresLicenseValidator
    {
        private readonly LicenseValidatorService _validator;

        public AresLicenseValidator()
        {
            _validator = new LicenseValidatorService();
        }

        public bool ValidateLicense()
        {
            return _validator.ValidateLicense();
        }

        public string GetLicenseInfo()
        {
            LicenseData licenseData;
            return _validator.GetLicenseInfo(out licenseData);
        }

        public string GetLastError()
        {
            return _validator.LastError ?? "";
        }

        public string GetCurrentUser()
        {
            return _validator.GetCurrentUser();
        }

        public int GetAuthorizedUserCount()
        {
            try
            {
                LicenseData licenseData;
                _validator.GetLicenseInfo(out licenseData);
                return licenseData?.MaxUsers ?? 0;
            }
            catch
            {
                return 0;
            }
        }
    }
}