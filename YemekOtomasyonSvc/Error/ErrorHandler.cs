using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YemekOtomasyonSvc.Error
{
    #region "Hata Event"
    /// <summary>
    /// OnError
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    public delegate void ErrorOccured(object sender, ErrorOccuredArgs e);

    public class ErrorOccuredArgs : EventArgs
    {
        private String _method = "";
        private String _message = "";

        public ErrorOccuredArgs(String method, String message)
        {
            _method = method;
            _message = message;
        }

        public String Method
        {
            get
            {
                return _method;
            }
        }

        public String Message
        {
            get
            {
                return _message;
            }
        }
    }
    #endregion
}
