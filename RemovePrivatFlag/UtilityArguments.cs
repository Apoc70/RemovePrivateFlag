/*
Based on: http://dotnetfollower.com/wordpress/2012/03/c-simple-command-line-arguments-parser/
 */
using System;

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
				return GetValue("mailbox");
			}
		}
		
		public bool noConfirmation
		{
			get
			{
				return GetBoolValue("-noconfirmation");
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
		
		public bool Help
		{
			get
			{
				return GetBoolValue("-help");
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
				return GetBoolValue("-logonly");
			}
		}
		
		public UtilityArguments(string[] args) : base(args)
	    {
	    }	
	}
}
