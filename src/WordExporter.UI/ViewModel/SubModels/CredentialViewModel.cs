using GalaSoft.MvvmLight;
using System;
using System.Security;

namespace WordExporter.UI.ViewModel.SubModels
{
    public class CredentialViewModel : ViewModelBase
    {
        private String _userName;

        public String UserName
        {
            get
            {
                return _userName;
            }
            set
            {
                Set<String>(() => this.UserName, ref _userName, value);
            }
        }

        private SecureString _password;

        public SecureString Password
        {
            get
            {
                return _password;
            }
            set
            {
                Set<SecureString>(() => this.Password, ref _password, value);
            }
        }
    }
}
