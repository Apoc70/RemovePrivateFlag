﻿// RemovePrivateFlag
//
// Author: Torsten Schlopsnies, Thomas Stensitzki
//
// Based on: http://dotnetfollower.com/wordpress/2012/03/c-simple-command-line-arguments-parser/
//
// Published under MIT license

namespace RemovePrivateFlag
{
    /// <summary>
    /// Description of UtilityArguments.
    /// </summary>
    public class UtilityArguments : InputArguments
    {
        public string Mailbox
        {
            get
            {
                return GetValue("-mailbox");
            }
        }

        public bool noconfirmation
        {
            get
            {
                return GetSwitchValue("-noconfirmation");
            }
        }

        protected bool GetBoolValue(string key)
        {
            string adjustedKey;
            if (ContainsKey(key, out adjustedKey))
            {
                bool res;
                bool.TryParse(_parsedArguments[adjustedKey], out res);
                return res;
            }
            return false;
        }

        protected bool GetSwitchValue(string key)
        {
            string adjustedKey;
            if (ContainsKey(key, out adjustedKey))
            {
                return true;
            }
            return false;
        }

        public bool Help
        {
            get
            {
                return GetSwitchValue("-help");
            }
        }

        public string Foldername
        {
            get
            {
                return GetValue("-foldername");
            }
        }

        public bool LogOnly
        {
            get
            {
                return GetSwitchValue("-logonly");
            }
        }

        public bool IgnoreSSLErrors
        {
            get
            {
                return GetSwitchValue("-ignoresslerrors");
            }
        }

        public UtilityArguments(string[] args) : base(args)
        {
        }
    }
}