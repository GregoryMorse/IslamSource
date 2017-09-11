using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Windows.UI.Xaml.Controls;
using Xamarin.Forms;
using Xamarin.Forms.Platform.UWP;
using Xamarin.Forms.Xaml;
//[assembly: XamlCompilation(XamlCompilationOptions.Compile)]

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("ISQV.Xam.UWP")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("ISQV.Xam.UWP")]
[assembly: AssemblyCopyright("Copyright ©  2015")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: ComVisible(false)]

[assembly: ExportRenderer(typeof(ISQV.Xam.UWP.CustomProgressRing), typeof(ISQV.Xam.UWP.CustomProgressRingRenderer))]
namespace ISQV.Xam.UWP
{
    public class CustomProgressRingRenderer : ViewRenderer<CustomProgressRing, ProgressRing>
    {
        ProgressRing ring;
        protected override void OnElementChanged(ElementChangedEventArgs<CustomProgressRing> e)
        {
            base.OnElementChanged(e);
            if (Control == null)
            {
                ring = new ProgressRing();
                ring.IsActive = true;
                ring.Visibility = Windows.UI.Xaml.Visibility.Visible;
                ring.IsEnabled = true;
                SetNativeControl(ring);
            }
        }
    }
    public class CustomProgressRing : View
    {
    }
}