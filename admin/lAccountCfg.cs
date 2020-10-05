using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sipek.Common;

namespace admin
{
    

    class rc_AccountCfg : IAccount
    {
        string phone_number = Properties.Settings.Default.phone_number;
        string phone_name = Properties.Settings.Default.phone_name;
        string phone_server = Properties.Settings.Default.phone_server;
        string phone_pass = Properties.Settings.Default.phone_pass;

        #region IAccount Members

        public string AccountName
        {
            get
            {
                return phone_number;
            }
            set
            {
            }
        }

        public string DisplayName
        {
            get
            {
                return phone_number;
            }
            set
            {
            }
        }

        public string DomainName
        {
            get
            {
                return "*";
            }
            set
            {
            }
        }

        public bool Enabled
        {
            get
            {
                return true;
            }
            set
            {
            }
        }

        public string HostName
        {
            get
            {
                return phone_server;
            }
            set
            {
            }
        }

        public string Id
        {
            get
            {
                return phone_number;
            }
            set
            {
            }
        }

        public int Index
        {
            get
            {
                return 0;
            }
            set
            {
            }
        }

        public string Password
        {
            get
            {
                return phone_pass;
            }
            set
            {
            }
        }

        public string ProxyAddress
        {
            get
            {
                return "";
            }
            set
            {
            }
        }

        public int RegState
        {
            get
            {
                return 0;
            }
            set
            {
            }
        }

        public ETransportMode TransportMode
        {
            get
            {
                return ETransportMode.TM_UDP;
            }
            set
            {
            }
        }

        public string UserName
        {
            get
            {
                return phone_number;
            }
            set
            {
            }
        }

        #endregion
    } // class rc_AccountCfg

} // namespace iLLi.VOIP