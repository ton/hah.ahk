// Author: Markus Scholtes, 2021
// Version 1.9, 2021-10-08
// Version for Windows 10 1809 to 21H1
// Compile with:
// C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe VirtualDesktop.cs

using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Reflection;

[assembly:AssemblyTitle("Command line tool to manage virtual desktops")]
[assembly:AssemblyDescription("Command line tool to manage virtual desktops")]
[assembly:AssemblyConfiguration("")]
[assembly:AssemblyCompany("MS")]
[assembly:AssemblyProduct("VirtualDesktop")]
[assembly:AssemblyCopyright("© Markus Scholtes 2021")]
[assembly:AssemblyTrademark("")]
[assembly:AssemblyCulture("")]
[assembly:AssemblyVersion("1.9.0.0")]
[assembly:AssemblyFileVersion("1.9.0.0")]

// Based on http://stackoverflow.com/a/32417530, Windows 10 SDK, github project Grabacr07/VirtualDesktop and own research

namespace VirtualDesktop
{
	#region COM API
	internal static class Guids
	{
		public static readonly Guid CLSID_ImmersiveShell = new Guid("C2F03A33-21F5-47FA-B4BB-156362A2F239");
		public static readonly Guid CLSID_VirtualDesktopManagerInternal = new Guid("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");
	}

  [ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIInspectable)]
	[Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
	internal interface IApplicationView
	{
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")]
	internal interface IVirtualDesktop
	{
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("F31574D6-B682-4CDC-BD56-1827860ABEC6")]
	internal interface IVirtualDesktopManagerInternal
	{
		int GetCount();
		void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
		bool CanViewMoveDesktops(IApplicationView view);
		IVirtualDesktop GetCurrentDesktop();
		void GetDesktops(out IObjectArray desktops);
		[PreserveSig]
		int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
		void SwitchDesktop(IVirtualDesktop desktop);
		IVirtualDesktop CreateDesktop();
		void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
		IVirtualDesktop FindDesktop(ref Guid desktopid);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
	internal interface IVirtualDesktopManager
	{
		bool IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
		Guid GetWindowDesktopId(IntPtr topLevelWindow);
		void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
	internal interface IObjectArray
	{
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
	internal interface IServiceProvider10
	{
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object QueryService(ref Guid service, ref Guid riid);
	}
	#endregion

	#region COM wrapper
	internal static class DesktopManager
	{
		static DesktopManager()
		{
			var shell = (IServiceProvider10)Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_ImmersiveShell));
			VirtualDesktopManagerInternal = (IVirtualDesktopManagerInternal)shell.QueryService(Guids.CLSID_VirtualDesktopManagerInternal, typeof(IVirtualDesktopManagerInternal).GUID);
		}

		internal static IVirtualDesktopManagerInternal VirtualDesktopManagerInternal;
	}
	#endregion

	#region public interface
	public class Desktop
	{
		public static void Left()
		{ // return desktop at the left of this one, null if none
			 IVirtualDesktop current = DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop();

			IVirtualDesktop left;
			if (DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(current, 3, out left) == 0) // 3 = LeftDirection
			  DesktopManager.VirtualDesktopManagerInternal.SwitchDesktop(left);
		}

		public static void Right()
		{ 
	    IVirtualDesktop current = DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop();

			IVirtualDesktop right;
			if (DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(current, 4, out right) == 0) // 4 = RightDirection
			  DesktopManager.VirtualDesktopManagerInternal.SwitchDesktop(right);
		}
	}
	#endregion
}

namespace VDeskTool
{
	static class Program
	{
		static int Main(string[] args)
		{
			try
			{
				if (args.Length == 0)
					Console.WriteLine("VirtualDesktop.exe left|right");
				else if (args.Length == 1)
				{
					if (args[0] == "left") 
						VirtualDesktop.Desktop.Left();
					if (args[0] == "right")
						VirtualDesktop.Desktop.Right();
				}
			}
			catch
			{
				return -1;
			}
			return 0;
		}
	}
}
