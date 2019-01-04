using System;
using System.Collections.Generic;
using System.Globalization;
using System.Data;
using System.Linq;
using System.Text;

namespace YemekOtomasyonSvc.Misc
{
    public static class Settings
    {
        private static Boolean _admin = false;

        private static String _connStr1630 = "";
        private static String _connStrLocal = "Server=localhost;Database=***;User Id=***;Password=***;";

        private static String startPath = @"D:\Turnike\Turnike Kontrol\YemekhaneKontrol.exe";
        private static Int32 _yemekhaneId;

        private static DataTable _ayarlar;
        private static DataTable _turnikeler;

        private static Decimal _azamiBakiye = 0;

        private static CultureInfo _bolge;

        public static Boolean AdminMod
        {
            get
            {
                return _admin;
            }
            set
            {
                _admin = value;
            }
        }

        public static String ConnectionStr1630
        {
            get
            {
                return _connStr1630;
            }
            set
            {
                _connStr1630 = value;
            }
        }

        public static String ConnectionStrLocal
        {
            get
            {
                return _connStrLocal;
            }
            set
            {
                _connStrLocal = value;
            }
        }

        public static String StartPath
        {
            get
            {
                return startPath;
            }
            set
            {
                startPath = value;
            }
        }

        public static Int32 YemekhaneID
        {
            get
            {
                return _yemekhaneId;
            }
            set
            {
                _yemekhaneId = value;
            }
        }

        public static DataTable Ayarlar
        {
            get
            {
                return _ayarlar;
            }
            set
            {
                _ayarlar = value;
            }
        }

        public static DataTable Turnikeler
        {
            get
            {
                return _turnikeler;
            }
            set
            {
                _turnikeler = value;
            }
        }

        public static Decimal AzamiBakiye
        {
            get
            {
                return _azamiBakiye;
            }
            set
            {
                _azamiBakiye = value;
            }
        }

        public static CultureInfo Bolge
        {
            get
            {
                return _bolge;
            }
            set
            {
                _bolge = value;
            }
        }
    }
}
