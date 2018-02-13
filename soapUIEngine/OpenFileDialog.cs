using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace Inflectra.RemoteLaunch.Engines.soapUI
{
	public class OpenFileDialog : FileDialog
	{
		public bool? ShowDialog(Window owner)
		{
			NativeMethods.OpenFileName ofn = ToOfn(owner);
			if (NativeMethods.GetOpenFileName(ofn))
			{
				FromOfn(ofn);
				return true;
			}
			else
			{
				FreeOfn(ofn);
				return false;
			}
		}
	}
}
