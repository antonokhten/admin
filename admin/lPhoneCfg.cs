using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sipek.Common;

namespace admin
{
    internal class rc_PhoneCfg : IConfiguratorInterface
    {
        List<IAccount> v_slAccList = new List<IAccount>();

        internal rc_PhoneCfg()
        {
            v_slAccList.Add(new rc_AccountCfg());
        }

        #region IConfiguratorInterface Members

        public bool AAFlag
        {
            get
            {
                return false;
            }
            set { }
        }

        public List<IAccount> Accounts
        {
            get
            {
                return v_slAccList;
            }
        }

        public bool CFBFlag
        {
            get
            {
                return false;
            }
            set
            {
            }
        }

        public string CFBNumber
        {
            get
            {
                return "";
            }
            set
            {
            }
        }

        public bool CFNRFlag
        {
            get
            {
                return false;
            }
            set
            {
            }
        }

        public string CFNRNumber
        {
            get
            {
                return "";
            }
            set
            {
            }
        }

        public bool CFUFlag
        {
            get
            {
                return false;
            }
            set
            {
            }
        }

        public string CFUNumber
        {
            get
            {
                return "";
            }
            set
            {
            }
        }

        public List<String> CodecList
        {
            get
            {
                List<String> slCodecs = new List<String>();
                slCodecs.Add("PCMA");
                return slCodecs;
            }
            set
            {
            }
        }

        public bool DNDFlag
        {
            get
            {
                return false;
            }
            set
            {
            }
        }

        public int DefaultAccountIndex
        {
            get
            {
                return 0;
            }
        }

        public bool IsNull
        {
            get
            {
                return false;
            }
        }

        public bool PublishEnabled
        {
            get
            {
                return false;
            }
            set
            {
            }
        }

        public int SIPPort
        {
            get
            {
                return 5060;
            }
            set
            {
            }
        }

        public void Save()
        {
            // TODO
        }

        #endregion
    } // class rc_PhoneCfg

} // namespace iLLi.VOIP