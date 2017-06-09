using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using IslamResources.My.Resources;

#if WINDOWS_APP || WINDOWS_PHONE_APP
using Windows.Foundation;
public class WindowsRTFileIO : XMLRender.PortableFileIO
{
    public async System.Threading.Tasks.Task<string[]> GetDirectoryFiles(string Path)
    {
        //System.Threading.Tasks.Task<Windows.Storage.StorageFolder> t = Windows.Storage.StorageFolder.GetFolderFromPathAsync(Path).AsTask();
        //t.Wait();
        Windows.Storage.StorageFolder folder = /*t.Result;*/ await Windows.Storage.StorageFolder.GetFolderFromPathAsync(Path);
        //System.Threading.Tasks.Task<IReadOnlyList<Windows.Storage.StorageFile>> tn = folder.GetFilesAsync().AsTask();
        //t.Wait();
        IReadOnlyList<Windows.Storage.StorageFile> files = /*tn.Result;*/ await folder.GetFilesAsync();
        return new List<string>(files.Select(file => file.Name)).ToArray();
    }
    public async System.Threading.Tasks.Task<Stream> LoadStream(string FilePath)
    {
        Windows.ApplicationModel.Resources.Core.ResourceCandidate rc = null;
        if (IslamSourceQuranViewer.App._resourceContext == null && Windows.UI.Xaml.Window.Current != null && Windows.UI.Xaml.Window.Current.CoreWindow != null) { IslamSourceQuranViewer.App._resourceContext = Windows.ApplicationModel.Resources.Core.ResourceContext.GetForCurrentView(); } else { IslamSourceQuranViewer.App._resourceContext = Windows.ApplicationModel.Resources.Core.ResourceContext.GetForViewIndependentUse(); }
        if (IslamSourceQuranViewer.App._resourceContext != null) { rc = Windows.ApplicationModel.Resources.Core.ResourceManager.Current.MainResourceMap.GetValue(FilePath.Replace(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, "ms-resource:///Files").Replace("\\", "/"), IslamSourceQuranViewer.App._resourceContext); }
        //System.Threading.Tasks.Task<Windows.Storage.StorageFile> t;
        //if (rc != null && rc.IsMatch)
        //{
        //    t = rc.GetValueAsFileAsync().AsTask();
        //}
        //else
        //{
        //    t = Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath).AsTask();
        //}
        //t.Wait();
        Windows.Storage.StorageFile file = /*t.Result;*/ ((rc != null && rc.IsMatch) ? await rc.GetValueAsFileAsync() : await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath));
        //System.Threading.Tasks.Task<Stream> tn = file.OpenStreamForReadAsync();
        //tn.Wait();
        Stream Stream = /*tn.Result;*/ await file.OpenStreamForReadAsync();
        return Stream;
    }
    public async System.Threading.Tasks.Task SaveStream(string FilePath, Stream Stream)
    {
        //System.Threading.Tasks.Task<Windows.Storage.StorageFolder> td = Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(FilePath)).AsTask();
        //td.Wait();
        //System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = td.Result.CreateFileAsync(System.IO.Path.GetFileName(FilePath), Windows.Storage.CreationCollisionOption.ReplaceExisting).AsTask();
        //t.Wait();
        Windows.Storage.StorageFile file = /*t.Result;*/ await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(FilePath))).CreateFileAsync(System.IO.Path.GetFileName(FilePath), Windows.Storage.CreationCollisionOption.ReplaceExisting);
        //System.Threading.Tasks.Task<Stream> tn = file.OpenStreamForWriteAsync();
        //tn.Wait();
        Stream File = /*tn.Result;*/ await file.OpenStreamForWriteAsync();
        File.Seek(0, SeekOrigin.Begin);
        byte[] Bytes = new byte[4096];
        int Read;
        Stream.Seek(0, SeekOrigin.Begin);
        Read = Stream.Read(Bytes, 0, Bytes.Length);
        while (Read != 0)
        {
            File.Write(Bytes, 0, Read);
            Read = Stream.Read(Bytes, 0, Bytes.Length);
        }
        File.Dispose();
    }
    public string CombinePath(params string[] Paths)
    {
        return System.IO.Path.Combine(Paths);
    }
    public async System.Threading.Tasks.Task DeleteFile(string FilePath)
    {
        //System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath).AsTask();
        //t.Wait();
        Windows.Storage.StorageFile file = /*t.Result;*/ await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        //file.DeleteAsync().AsTask().Wait();
        await file.DeleteAsync();
    }
    public async System.Threading.Tasks.Task<bool> PathExists(string Path)
    {
        //System.Threading.Tasks.Task<Windows.Storage.StorageFolder> t = Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path)).AsTask();
        //t.Wait();
#if WINDOWS_PHONE_APP
        //System.Threading.Tasks.Task<IReadOnlyList<Windows.Storage.IStorageItem>> files = t.Result.GetItemsAsync().AsTask();
        //files.Wait();
        //return files.Result.FirstOrDefault(p => p.Name.Equals(System.IO.Path.GetFileName(Path), StringComparison.OrdinalIgnoreCase)) != null;
        return (await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).GetItemsAsync()).FirstOrDefault(p => p.Name.Equals(System.IO.Path.GetFileName(Path), StringComparison.OrdinalIgnoreCase)) != null;
#else
        //System.Threading.Tasks.Task<Windows.Storage.IStorageItem> tn = t.Result.TryGetItemAsync(System.IO.Path.GetFileName(Path)).AsTask();
        //tn.Wait();
        //return tn.Result != null;
        return (await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).TryGetItemAsync(System.IO.Path.GetFileName(Path))) != null;
#endif
    }
    public async System.Threading.Tasks.Task CreateDirectory(string Path)
    {
        //System.Threading.Tasks.Task<Windows.Storage.StorageFolder> t = Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path)).AsTask();
        //t.Wait();
        //t.Result.CreateFolderAsync(System.IO.Path.GetFileName(Path), Windows.Storage.CreationCollisionOption.OpenIfExists).AsTask().Wait();
        await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).CreateFolderAsync(System.IO.Path.GetFileName(Path), Windows.Storage.CreationCollisionOption.OpenIfExists);
    }
    public async System.Threading.Tasks.Task<DateTime> PathGetLastWriteTimeUtc(string Path)
    {
        //System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = Windows.Storage.StorageFile.GetFileFromPathAsync(Path).AsTask();
        //t.Wait();
        Windows.Storage.StorageFile file = /*t.Result;*/ await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        //System.Threading.Tasks.Task<Windows.Storage.FileProperties.BasicProperties> tn = file.GetBasicPropertiesAsync().AsTask();
        //tn.Wait();
        //return tn.Result.DateModified.UtcDateTime;
        return (await file.GetBasicPropertiesAsync()).DateModified.UtcDateTime;
    }
    public async System.Threading.Tasks.Task PathSetLastWriteTimeUtc(string Path, DateTime Time)
    {
        //System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = Windows.Storage.StorageFile.GetFileFromPathAsync(Path).AsTask();
        //t.Wait();
        Windows.Storage.StorageFile file = /*t.Result;*/ await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        //System.Threading.Tasks.Task<Windows.Storage.StorageStreamTransaction> tn = file.OpenTransactedWriteAsync().AsTask();
        //tn.Wait();
        //tn.Result.CommitAsync().AsTask().Wait();
        await (await file.OpenTransactedWriteAsync()).CommitAsync();
    }
}
public class WindowsRTSettings : XMLRender.PortableSettings
{
    public string CacheDirectory
    {
        get
        {
            //Windows.Storage.ApplicationData.Current.LocalFolder.InstalledLocation;
            //Windows.ApplicationModel.Package.Current.InstalledLocation.Path;
            return Windows.Storage.ApplicationData.Current.TemporaryFolder.Path;
        }
    }
    public KeyValuePair<string, string[]>[] Resources
    {
        get
        {
            return (new List<KeyValuePair<string, string[]>>(System.Linq.Enumerable.Select("HostPageUtility=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(';'), Str => new KeyValuePair<string, string[]>(Str.Split('=')[0], Str.Split('=')[1].Split(','))))).ToArray();
        }
    }
    public string[] FuncLibs
    {
        get
        {
            return new string[] { "IslamMetadata" };
        }
    }
    public string GetTemplatePath(string Selector)
    {
        return GetFilePath("metadata\\IslamSource.xml");
    }
    public string GetFilePath(string Path)
    {
        return System.IO.Path.Combine(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, Path);
    }
    public string GetUName(char Character)
    {
        return "";
    }
    public string GetResourceString(string baseName, string resourceKey)
    {
        return Windows.ApplicationModel.Resources.ResourceLoader.GetForViewIndependentUse(baseName + ".Resources").GetString(resourceKey);
    }
}
#elif WINDOWS_PHONE
public class WindowsRTXamFileIO : XMLRender.PortableFileIO
{
    public async System.Threading.Tasks.Task<string[]> GetDirectoryFiles(string Path)
    {
        Windows.Storage.StorageFolder folder = await Windows.Storage.StorageFolder.GetFolderFromPathAsync(Path);
        IReadOnlyList<Windows.Storage.StorageFile> files = await folder.GetFilesAsync();
        return new List<string>(files.Select(file => file.Name)).ToArray();
    }
    public async System.Threading.Tasks.Task<Stream> LoadStream(string FilePath)
    {
        Stream rc = null;
        rc = IslamSourceQuranViewer.Xam.WinPhone.Resources.AppResources.ResourceManager.GetStream(FilePath.Replace(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, "ms-resource:///Files").Replace("\\", "/"));
        if (rc != null)
        {
            return rc;
        }
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        Stream Stream = await file.OpenStreamForReadAsync();
        return Stream;
    }
    public async System.Threading.Tasks.Task SaveStream(string FilePath, Stream Stream)
    {
        Windows.Storage.StorageFile file = await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(FilePath))).CreateFileAsync(System.IO.Path.GetFileName(FilePath), Windows.Storage.CreationCollisionOption.ReplaceExisting);
        Stream File = await file.OpenStreamForWriteAsync();
        File.Seek(0, SeekOrigin.Begin);
        byte[] Bytes = new byte[4096];
        int Read;
        Stream.Seek(0, SeekOrigin.Begin);
        Read = Stream.Read(Bytes, 0, Bytes.Length);
        while (Read != 0)
        {
            File.Write(Bytes, 0, Read);
            Read = Stream.Read(Bytes, 0, Bytes.Length);
        }
        File.Dispose();
    }
    public string CombinePath(params string[] Paths)
    {
        return System.IO.Path.Combine(Paths);
    }
    public async System.Threading.Tasks.Task DeleteFile(string FilePath)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        await file.DeleteAsync();
    }
    public async System.Threading.Tasks.Task<bool> PathExists(string Path)
    {
        return (await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).GetItemsAsync()).FirstOrDefault(p => p.Name == Path) != null;
    }
    public async System.Threading.Tasks.Task CreateDirectory(string Path)
    {
        await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).CreateFolderAsync(System.IO.Path.GetFileName(Path), Windows.Storage.CreationCollisionOption.OpenIfExists);
    }
    public async System.Threading.Tasks.Task<DateTime> PathGetLastWriteTimeUtc(string Path)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        return (await file.GetBasicPropertiesAsync()).DateModified.UtcDateTime;
    }
    public async System.Threading.Tasks.Task PathSetLastWriteTimeUtc(string Path, DateTime Time)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        await(await file.OpenTransactedWriteAsync()).CommitAsync();
    }
}
public class WindowsRTXamSettings : XMLRender.PortableSettings
{
    public string CacheDirectory
    {
        get
        {
            return System.IO.Path.GetTempPath();
        }
    }
    public KeyValuePair<string, string[]>[] Resources
    {
        get
        {
            return (new List<KeyValuePair<string, string[]>>(System.Linq.Enumerable.Select("HostPageUtility=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(';'), Str => new KeyValuePair<string, string[]>(Str.Split('=')[0], Str.Split('=')[1].Split(','))))).ToArray();
        }
    }
    public string[] FuncLibs
    {
        get
        {
            return new string[] { "IslamMetadata" };
        }
    }
    public string GetTemplatePath(string Selector)
    {
        return GetFilePath("metadata\\IslamSource.xml");
    }
    public string GetFilePath(string Path)
    {
        return System.IO.Path.Combine(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, Path);
    }
    public string GetUName(char Character)
    {
        return "";
    }
    public string GetResourceString(string baseName, string resourceKey)
    {
        return new System.Resources.ResourceManager(baseName + ".Resources", System.Reflection.Assembly.Load(baseName)).GetString(resourceKey, System.Threading.Thread.CurrentThread.CurrentUICulture);
    }
}
#else
public class AndroidiOSFileIO : XMLRender.PortableFileIO
{
    public async System.Threading.Tasks.Task<string[]> GetDirectoryFiles(string Path)
    {
        return System.IO.Directory.GetFiles(Path);
    }
    public async System.Threading.Tasks.Task<Stream> LoadStream(string FilePath)
    {
#if __ANDROID__
        return ((IslamSourceQuranViewer.Xam.Droid.MainActivity)global::Xamarin.Forms.Forms.Context).Assets.Open(FilePath.Trim('/'));
#endif
#if __IOS__
        return File.Open(FilePath, FileMode.Open, FileAccess.Read);
#endif
    }
    public async System.Threading.Tasks.Task SaveStream(string FilePath, Stream Stream)
    {
        FileStream File = System.IO.File.Open(FilePath, FileMode.Create, FileAccess.Write);
        File.Seek(0, SeekOrigin.Begin);
        byte[] Bytes = new byte[4096];
        int Read;
        Stream.Seek(0, SeekOrigin.Begin);
        Read = Stream.Read(Bytes, 0, Bytes.Length);
        while (Read != 0)
        {
            File.Write(Bytes, 0, Read);
            Read = Stream.Read(Bytes, 0, Bytes.Length);
        }
        File.Dispose();
    }
    public string CombinePath(params string[] Paths)
    {
        return System.IO.Path.Combine(Paths);
    }
    public async System.Threading.Tasks.Task DeleteFile(string FilePath)
    {
        File.Delete(FilePath);
    }
    public async System.Threading.Tasks.Task<bool> PathExists(string Path)
    {
        return Directory.Exists(Path);
    }
    public async System.Threading.Tasks.Task CreateDirectory(string Path)
    {
        Directory.CreateDirectory(Path);
    }
    public async System.Threading.Tasks.Task<DateTime> PathGetLastWriteTimeUtc(string Path)
    {
        return System.IO.File.GetLastWriteTimeUtc(Path);
    }
    public async System.Threading.Tasks.Task PathSetLastWriteTimeUtc(string Path, DateTime Time)
    {
        System.IO.File.SetLastWriteTimeUtc(Path, Time);
    }
}
public class AndroidiOSSettings : XMLRender.PortableSettings
{
    public string CacheDirectory
    {
        get
        {
#if __ANDROID__
            return ((IslamSourceQuranViewer.Xam.Droid.MainActivity)global::Xamarin.Forms.Forms.Context).CacheDir.Path;
#endif
#if __IOS__
            return string.Empty;
#endif
        }
    }
    public KeyValuePair<string, string[]>[] Resources
    {
        get
        {
            return (new List<KeyValuePair<string, string[]>>(System.Linq.Enumerable.Select("HostPageUtility=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(';'), Str => new KeyValuePair<string, string[]>(Str.Split('=')[0], Str.Split('=')[1].Split(','))))).ToArray();
        }
    }
    public string[] FuncLibs
    {
        get
        {
            return new string[] { "IslamMetadata" };
        }
    }
    public string GetTemplatePath(string Selector)
    {
        return GetFilePath("metadata/IslamSource.xml");
    }
    public string GetFilePath(string Path)
    {
        return System.IO.Path.Combine(Directory.GetCurrentDirectory(), Path);
    }
    public string GetUName(char Character)
    {
        return "";
    }
    public string GetResourceString(string baseName, string resourceKey)
    {
        return new System.Resources.ResourceManager(baseName + ".Resources", System.Reflection.Assembly.Load(baseName)).GetString(resourceKey, System.Threading.Thread.CurrentThread.CurrentUICulture);
    }
}
#endif

namespace IslamSourceQuranViewer
{
#if __IOS__

    using System.Drawing;
    using Foundation;
    using UIKit;
    public static class TextShaping
    {
        public static void Cleanup(int AllNormArb)
        {
        }
        public static double CalculateWidth(string text, bool IsArabic, float maxWidth, float maxHeight)
        {
            string FontFamily = "";
            double FontSize = 12.0;
            double width = TextMeterImplementation.MeasureTextSize(text, maxWidth, FontSize, FontFamily).Width;
            return width;
        }
        public static short[] GetWordDiacriticClusters(string Str, string useFont, float fontSize, bool IsRTL)
        {
            return new short[] { };
        }
        public static List<string> GetFontList()
        {
            List<string> fontList = new List<string>();
            return fontList;
        }
        public static string GetAppLanguage()
        {
            return string.Empty;
        }
        public static void SetAppLanguage(string Lang)
        {
        }
        public static string GetSelectedAppLanguage()
        {
            return string.Empty;
        }
        public static List<string> GetAppLanguageList()
        {
            return new List<string>();
        }
    }

    public static class TextMeterImplementation
    {
        public static Xamarin.Forms.Size MeasureTextSize(string text, double width,
            double fontSize, string fontName = null)
        {
            var nsText = new NSString(text);
            var boundSize = new SizeF((float)width, float.MaxValue);
            var options = NSStringDrawingOptions.UsesFontLeading |
                NSStringDrawingOptions.UsesLineFragmentOrigin;

            if (fontName == null)
            {
                fontName = "HelveticaNeue";
            }

            var attributes = new UIStringAttributes
            {
                Font = UIFont.FromName(fontName, (float)fontSize)
            };

            var sizeF = nsText.GetBoundingRect(boundSize, options, attributes, null).Size;

            return new Xamarin.Forms.Size((double)sizeF.Width, (double)sizeF.Height);
        }
    }

#elif __ANDROID__

using Android.Widget;
using Android.Util;
using Android.Views;
using Android.Graphics;
    public static class TextShaping
    {
        public static void Cleanup(int AllNormArb)
        {
        }
        public static double CalculateWidth(string text, bool IsArabic, float maxWidth, float maxHeight)
        {
            string FontFamily = "";
            double FontSize = 12.0;
            double width = TextMeterImplementation.MeasureTextSize(text, maxWidth, FontSize, FontFamily).Width;
            return width;
        }
        public static short[] GetWordDiacriticClusters(string Str, string useFont, float fontSize, bool IsRTL)
        {
            return new short[] {};
        }
        public static List<string> GetFontList()
        {
            List<string> fontList = new List<string>();
            return fontList;
        }
        public static string GetAppLanguage()
        {
            return string.Empty;
        }
        public static void SetAppLanguage(string Lang)
        {
        }
        public static string GetSelectedAppLanguage()
        {
            return string.Empty;
        }
        public static List<string> GetAppLanguageList()
        {
            return new List<string>();
        }
    }
	public static class TextMeterImplementation
	{
		private static Typeface textTypeface;

		public static Xamarin.Forms.Size MeasureTextSize(string text, double width,
			double fontSize, string fontName = null)
		{
			var textView = new TextView(global::Android.App.Application.Context);
			textView.Typeface = GetTypeface(fontName);
			textView.SetText(text, TextView.BufferType.Normal);
			textView.SetTextSize(ComplexUnitType.Px, (float)fontSize);

			int widthMeasureSpec = Android.Views.View.MeasureSpec.MakeMeasureSpec(
				(int)width, MeasureSpecMode.AtMost);
			int heightMeasureSpec = Android.Views.View.MeasureSpec.MakeMeasureSpec(
				0, MeasureSpecMode.Unspecified);

			textView.Measure(widthMeasureSpec, heightMeasureSpec);

			return new Xamarin.Forms.Size((double)textView.MeasuredWidth,
				(double)textView.MeasuredHeight);
		}

		private static Typeface GetTypeface(string fontName)
		{
			if (fontName == null)
			{
				return Typeface.Default;
			}

			if (textTypeface == null)
			{
				textTypeface = Typeface.Create(fontName, TypefaceStyle.Normal);
			}

			return textTypeface;
		}
	}
#elif WINDOWS_APP || WINDOWS_PHONE_APP
    public static class TextShaping
    {
        public static void Cleanup(int AllNormArb)
        {
            if (AllNormArb == 0)
            {
                _DWFeatureArray = null;
                if (_DWFontFace != null) _DWFontFace.Dispose();
                _DWFontFace = null;
                if (_DWFont != null) _DWFont.Dispose();
                _DWFont = null;
                if (_DWAnalyzer != null) _DWAnalyzer.Dispose();
                _DWAnalyzer = null;
            }
            if (AllNormArb != 2)
            {
                if (_DWNormalFormat != null) _DWNormalFormat.Dispose();
                _DWNormalFormat = null;
            }
            if (AllNormArb != 1)
            {
                if (_DWArabicFormat != null) _DWArabicFormat.Dispose();
                _DWArabicFormat = null;
            }
            if (AllNormArb == 0)
            {
                if (_DWFactory != null) _DWFactory.Dispose();
                _DWFactory = null;
            }
        }
        private static SharpDX.DirectWrite.Factory _DWFactory;
        public static SharpDX.DirectWrite.Factory DWFactory { get { if (_DWFactory == null) _DWFactory = new SharpDX.DirectWrite.Factory(); return _DWFactory; } }
        private static SharpDX.DirectWrite.TextFormat _DWArabicFormat;
        public static SharpDX.DirectWrite.TextFormat DWArabicFormat { get { if (_DWArabicFormat == null) _DWArabicFormat = new SharpDX.DirectWrite.TextFormat(DWFactory, AppSettings.strSelectedFont, (float)AppSettings.dFontSize); return _DWArabicFormat; } set { if (_DWArabicFormat != null) _DWArabicFormat.Dispose(); _DWArabicFormat = value; } }
        private static SharpDX.DirectWrite.TextFormat _DWNormalFormat;
        public static SharpDX.DirectWrite.TextFormat DWNormalFormat { get { if (_DWNormalFormat == null) _DWNormalFormat = new SharpDX.DirectWrite.TextFormat(DWFactory, AppSettings.strOtherSelectedFont, (float)AppSettings.dOtherFontSize); return _DWNormalFormat; } set { if (_DWNormalFormat != null) _DWNormalFormat.Dispose(); _DWNormalFormat = value; } }
        private static SharpDX.DirectWrite.TextAnalyzer _DWAnalyzer;
        public static SharpDX.DirectWrite.TextAnalyzer DWAnalyzer { get { if (_DWAnalyzer == null) _DWAnalyzer = new SharpDX.DirectWrite.TextAnalyzer(DWFactory); return _DWAnalyzer; } }
        private static SharpDX.DirectWrite.FontFeature[] _DWFeatureArray;
        public static SharpDX.DirectWrite.FontFeature[] DWFeatureArray { get { if (_DWFeatureArray == null) _DWFeatureArray = new SharpDX.DirectWrite.FontFeature[] { new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.GlyphCompositionDecomposition, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.DiscretionaryLigatures, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StandardLigatures, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.ContextualAlternates, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StylisticSet1, 1) }; return _DWFeatureArray; } }
        private static SharpDX.DirectWrite.Font _DWFont;
        private static SharpDX.DirectWrite.FontFace _DWFontFace;
        public static SharpDX.DirectWrite.FontFace DWFontFace
        {
            get
            {
                //LOGFONT lf = new LOGFONT();
                //lf.lfFaceName = useFont;
                //SharpDX.Direct2D1.Factory fact2d = new SharpDX.Direct2D1.Factory();
                //float pointSize = fontSize * fact2d.DesktopDpi.Height / 72.0f;
                //fact2d.Dispose();
                //lf.lfHeight = (int)fontSize;
                //lf.lfQuality = 5; //clear type
                //SharpDX.DirectWrite.Font font = MyUIChanger.DWFactory.GdiInterop.FromLogFont(lf);
                if (_DWFontFace == null)
                {
                    int index;
                    DWArabicFormat.FontCollection.FindFamilyName(DWArabicFormat.FontFamilyName, out index);
                    _DWFont = DWArabicFormat.FontCollection.GetFontFamily(index).GetFirstMatchingFont(SharpDX.DirectWrite.FontWeight.Normal, SharpDX.DirectWrite.FontStretch.Normal, SharpDX.DirectWrite.FontStyle.Normal);
                    _DWFontFace = new SharpDX.DirectWrite.FontFace(_DWFont);
                    //fontFace.FaceType = SharpDX.DirectWrite.FontFaceType.
                }
                return _DWFontFace;
            }
        }
        public static double CalculateWidth(string text, bool IsArabic, float maxWidth, float maxHeight)
        {
            SharpDX.DirectWrite.TextLayout layout = new SharpDX.DirectWrite.TextLayout(DWFactory, text, IsArabic ? DWArabicFormat : DWNormalFormat, maxWidth, maxHeight);
            double width = layout.Metrics.WidthIncludingTrailingWhitespace + layout.Metrics.Left;
            layout.Dispose();
            return width;
        }
        public static List<string> GetFontList()
        {
            List<string> fontList = new List<string>();
            SharpDX.DirectWrite.FontCollection fontCollection = DWFactory.GetSystemFontCollection(false);
            for (int i = 0; i < fontCollection.FontFamilyCount; i++)
            {
                int index = 0;
                if (!fontCollection.GetFontFamily(i).FamilyNames.FindLocaleName(System.Globalization.CultureInfo.CurrentCulture.Name, out index))
                {
                    for (int j = 0; j < Windows.Globalization.ApplicationLanguages.Languages.Count; j++)
                    {
                        if (fontCollection.GetFontFamily(i).FamilyNames.FindLocaleName(Windows.Globalization.ApplicationLanguages.Languages[j], out index))
                        {
                            fontList.Add(fontCollection.GetFontFamily(i).FamilyNames.GetString(index));
                            break;
                        }
                    }

                }
                else { fontList.Add(fontCollection.GetFontFamily(i).FamilyNames.GetString(index)); }

            }
            return fontList;
        }
        public static string GetAppLanguage()
        {
            return (string)Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride;
        }
        public static void SetAppLanguage(string Lang)
        {
            Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = (Lang == Windows.Globalization.ApplicationLanguages.Languages.First() ? string.Empty : Lang);
            App._resourceContext = null;
            if (Windows.UI.Xaml.Window.Current != null && Windows.UI.Xaml.Window.Current.CoreWindow != null) Windows.ApplicationModel.Resources.Core.ResourceContext.GetForCurrentView().Reset();
            Windows.ApplicationModel.Resources.Core.ResourceContext.GetForViewIndependentUse().Reset();
            Windows.ApplicationModel.Resources.Core.ResourceContext.ResetGlobalQualifierValues();
            System.Globalization.CultureInfo.DefaultThreadCurrentCulture = new System.Globalization.CultureInfo(Lang == string.Empty ? Windows.Globalization.ApplicationLanguages.Languages.First() : Lang);
            System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo(Lang == string.Empty ? Windows.Globalization.ApplicationLanguages.Languages.First() : Lang);
            while (System.Globalization.CultureInfo.CurrentUICulture.Name != (Lang == string.Empty ? Windows.Globalization.ApplicationLanguages.Languages.First() : Lang))
            {
                System.Threading.Tasks.Task.Delay(100);
            }
        }
        public static string GetSelectedAppLanguage()
        {
            return new Windows.Globalization.Language(GetAppLanguage() == string.Empty ? Windows.Globalization.ApplicationLanguages.Languages.First() : GetAppLanguage()).DisplayName + " (" + (GetAppLanguage() == string.Empty ? Windows.Globalization.ApplicationLanguages.Languages.First() : GetAppLanguage()) + ")";
        }
        public static List<string> GetAppLanguageList()
        {
            return Windows.Globalization.ApplicationLanguages.ManifestLanguages.Select((Item) => new Windows.Globalization.Language(Item).DisplayName + " (" + Item + ")").ToList();
        }
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct LayoutInfo
        {
            public Rect Rect;
            public float Baseline;
            public int nChar;
            public List<List<List<LayoutInfo>>> Bounds;
            public LayoutInfo(Rect NewRect, float NewBaseline, int NewNChar, List<List<List<LayoutInfo>>> NewBounds)
            {
                this = new LayoutInfo();
                this.Rect = NewRect;
                this.Baseline = NewBaseline;
                this.nChar = NewNChar;
                this.Bounds = NewBounds;
            }
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct CharPosInfo
        {
            public int Index;
            public int Length;
            public float Width;
            public float PriorWidth;
            public float X;
            public float Y;
            public float Height;
        }

        const int ERROR_INSUFFICIENT_BUFFER = 122;

        public class TextSource : SharpDX.DirectWrite.TextAnalysisSource
        {
            // Fields
            public SharpDX.DirectWrite.Factory _Factory;
            private string _Str;
            private bool disposedValue;

            // Methods
            public TextSource(string Str, SharpDX.DirectWrite.Factory Factory)
            {
                this._Str = Str;
                this._Factory = Factory;
            }

            public void Dispose()
            {
                this.Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!this.disposedValue && disposing)
                {
                }
                this.disposedValue = true;
            }

            public string GetLocaleName(int textPosition, out int textLength)
            {
                textLength = _Str.Length - textPosition;
                return System.Globalization.CultureInfo.CurrentCulture.Name;
            }

            public SharpDX.DirectWrite.NumberSubstitution GetNumberSubstitution(int textPosition, out int textLength)
            {
                textLength = _Str.Length - textPosition;
                return new SharpDX.DirectWrite.NumberSubstitution(this._Factory, SharpDX.DirectWrite.NumberSubstitutionMethod.None, null, true);
            }

            public string GetTextAtPosition(int textPosition)
            {
                return this._Str.Substring(textPosition);
            }

            public string GetTextBeforePosition(int textPosition)
            {
                return this._Str.Substring(0x0, textPosition - 0x1);
            }

            // Properties
            public SharpDX.DirectWrite.ReadingDirection ReadingDirection
            {
                get
                {
                    return SharpDX.DirectWrite.ReadingDirection.RightToLeft;
                }
            }


            public System.IDisposable Shadow { get; set; }
        }

        public class TextSink : SharpDX.DirectWrite.TextAnalysisSink
        {
            public byte _explicitLevel;
            public SharpDX.DirectWrite.LineBreakpoint[] _lineBreakpoints;
            public SharpDX.DirectWrite.NumberSubstitution _numberSubstitution;
            public byte _resolvedLevel;
            public SharpDX.DirectWrite.ScriptAnalysis _scriptAnalysis;
            private bool disposedValue;

            public void Dispose()
            {
                this.Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!this.disposedValue && disposing)
                {
                }
                this.disposedValue = true;
            }

            public void SetBidiLevel(int textPosition, int textLength, byte explicitLevel, byte resolvedLevel)
            {
                this._explicitLevel = explicitLevel;
                this._resolvedLevel = resolvedLevel;
            }

            public void SetLineBreakpoints(int textPosition, int textLength, SharpDX.DirectWrite.LineBreakpoint[] lineBreakpoints)
            {
                this._lineBreakpoints = lineBreakpoints;
            }

            public void SetNumberSubstitution(int textPosition, int textLength, SharpDX.DirectWrite.NumberSubstitution numberSubstitution)
            {
                this._numberSubstitution = numberSubstitution;
            }

            public void SetScriptAnalysis(int textPosition, int textLength, SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis)
            {
                this._scriptAnalysis = scriptAnalysis;
            }

            public IDisposable Shadow { get; set; }
        }
        //LOGFONT struct
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        public class LOGFONT
        {
            public const int LF_FACESIZE = 32;
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public int lfWeight;
            public byte lfItalic;
            public byte lfUnderline;
            public byte lfStrikeOut;
            public byte lfCharSet;
            public byte lfOutPrecision;
            public byte lfClipPrecision;
            public byte lfQuality;
            public byte lfPitchAndFamily;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = LF_FACESIZE)]
            public string lfFaceName;
        }
        public static short[] GetWordDiacriticClusters(string Str, string useFont, float fontSize, bool IsRTL)
        {
            if (Str == string.Empty)
            {
                return null;
            }
            SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis;
            TextSink analysisSink = new TextSink();
            TextSource analysisSource = new TextSource(Str, DWFactory);
            DWAnalyzer.AnalyzeScript(analysisSource, 0, Str.Length, analysisSink);
            scriptAnalysis = analysisSink._scriptAnalysis;
            int maxGlyphCount = ((Str.Length * 3) / 2) + 0x10;
            short[] clusterMap = new short[(Str.Length - 1) + 1];
            SharpDX.DirectWrite.ShapingTextProperties[] textProps = new SharpDX.DirectWrite.ShapingTextProperties[(Str.Length - 1) + 1];
            short[] glyphIndices = new short[(maxGlyphCount - 1) + 1];
            SharpDX.DirectWrite.ShapingGlyphProperties[] glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
            int actualGlyphCount = 0;
            while (true)
            {
                try
                {
                    DWAnalyzer.GetGlyphs(Str, Str.Length, DWFontFace, false, IsRTL, scriptAnalysis, null, null, new SharpDX.DirectWrite.FontFeature[][] { DWFeatureArray }, new int[] { Str.Length }, maxGlyphCount, clusterMap, textProps, glyphIndices, glyphProps, out actualGlyphCount);
                    break;
                }
                catch (SharpDX.SharpDXException exception)
                {
                    if (exception.ResultCode == SharpDX.Result.GetResultFromWin32Error(0x7a))
                    {
                        maxGlyphCount *= 2;
                        glyphIndices = new short[(maxGlyphCount - 1) + 1];
                        glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
                    }
                }
            }
            Array.Resize(ref glyphIndices, (actualGlyphCount - 1) + 1);
            Array.Resize(ref glyphProps, (actualGlyphCount - 1) + 1);
            float[] glyphAdvances = new float[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.GlyphOffset[] glyphOffsets = new SharpDX.DirectWrite.GlyphOffset[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.FontFeature[][] features = new SharpDX.DirectWrite.FontFeature[][] { DWFeatureArray };
            int[] featureRangeLengths = new int[] { Str.Length };
            DWAnalyzer.GetGlyphPlacements(Str, clusterMap, textProps, Str.Length, glyphIndices, glyphProps, actualGlyphCount, DWFontFace, fontSize, false, IsRTL, scriptAnalysis, null, features, featureRangeLengths, glyphAdvances, glyphOffsets);
            analysisSource.Shadow.Dispose();
            analysisSink.Shadow.Dispose();
            analysisSource.Dispose();
            analysisSource._Factory = null;
            analysisSink.Dispose();
            return clusterMap;
        }
        public static Size GetWordDiacriticPositionsDWrite(string Str, string useFont, float fontSize, char[] Forms, bool IsRTL, ref float BaseLine, ref CharPosInfo[] Pos)
        {
            if (Str == string.Empty)
            {
                return new Size(0f, 0f);
            }
            SharpDX.DirectWrite.TextAnalyzer analyzer = new SharpDX.DirectWrite.TextAnalyzer(DWFactory);
            LOGFONT lf = new LOGFONT();
            lf.lfFaceName = useFont;
            float pointSize = fontSize * Windows.Graphics.Display.DisplayInformation.GetForCurrentView().RawDpiY / 72.0f;
            lf.lfHeight = (int)fontSize;
            lf.lfQuality = 5; //clear type
            SharpDX.DirectWrite.Font font = DWFactory.GdiInterop.FromLogFont(lf);
            SharpDX.DirectWrite.FontFace fontFace = new SharpDX.DirectWrite.FontFace(font);
            SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis = new SharpDX.DirectWrite.ScriptAnalysis();
            TextSink analysisSink = new TextSink();
            TextSource analysisSource = new TextSource(Str, DWFactory);
            analyzer.AnalyzeScript(analysisSource, 0, Str.Length, analysisSink);
            scriptAnalysis = analysisSink._scriptAnalysis;
            int maxGlyphCount = ((Str.Length * 3) / 2) + 0x10;
            short[] clusterMap = new short[(Str.Length - 1) + 1];
            SharpDX.DirectWrite.ShapingTextProperties[] textProps = new SharpDX.DirectWrite.ShapingTextProperties[(Str.Length - 1) + 1];
            short[] glyphIndices = new short[(maxGlyphCount - 1) + 1];
            SharpDX.DirectWrite.ShapingGlyphProperties[] glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
            int actualGlyphCount = 0;
            SharpDX.DirectWrite.FontFeature[] featureArray = new SharpDX.DirectWrite.FontFeature[] { new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.GlyphCompositionDecomposition, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.DiscretionaryLigatures, 0), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StandardLigatures, 0), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.ContextualAlternates, 0), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StylisticSet1, 0) };
            while (true)
            {
                try
                {
                    analyzer.GetGlyphs(Str, Str.Length, fontFace, false, IsRTL, scriptAnalysis, null, null, new SharpDX.DirectWrite.FontFeature[][] { featureArray }, new int[] { Str.Length }, maxGlyphCount, clusterMap, textProps, glyphIndices, glyphProps, out actualGlyphCount);
                    break;
                }
                catch (SharpDX.SharpDXException exception)
                {
                    if (exception.ResultCode == SharpDX.Result.GetResultFromWin32Error(0x7a))
                    {
                        maxGlyphCount *= 2;
                        glyphIndices = new short[(maxGlyphCount - 1) + 1];
                        glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
                    }
                }
            }
            Array.Resize(ref glyphIndices, (actualGlyphCount - 1) + 1);
            Array.Resize(ref glyphProps, (actualGlyphCount - 1) + 1);
            float[] glyphAdvances = new float[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.GlyphOffset[] glyphOffsets = new SharpDX.DirectWrite.GlyphOffset[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.FontFeature[][] features = new SharpDX.DirectWrite.FontFeature[][] { featureArray };
            int[] featureRangeLengths = new int[] { Str.Length };
            analyzer.GetGlyphPlacements(Str, clusterMap, textProps, Str.Length, glyphIndices, glyphProps, actualGlyphCount, fontFace, fontSize, false, IsRTL, scriptAnalysis, null, features, featureRangeLengths, glyphAdvances, glyphOffsets);
            List<CharPosInfo> list = new List<CharPosInfo>();
            float PriorWidth = 0f;
            int RunStart = 0;
            int RunRes = clusterMap[0];
            if (IsRTL)
            {
                XMLRender.ArabicData.LigatureInfo[] array = AppSettings.ArbData.GetLigatures(Str, false, Forms);
                for (int CharCount = 0; CharCount < clusterMap.Length - 1; CharCount++)
                {
                    int RunCount = 0;
                    for (int ResCount = clusterMap[CharCount]; ResCount <= ((CharCount == (clusterMap.Length - 1)) ? (actualGlyphCount - 1) : (clusterMap[CharCount + 1] - 1)); ResCount++)
                    {
                        if ((glyphAdvances[ResCount] == 0f) & ((clusterMap.Length <= (RunStart + RunCount)) || (clusterMap[RunStart] == clusterMap[RunStart + RunCount])))
                        {
                            int Index = Array.FindIndex<XMLRender.ArabicData.LigatureInfo>(array, (lig) => lig.Indexes[0] == RunStart + RunCount);
                            int LigLen = 1;
                            if (Index != -1)
                            {
                                while ((LigLen != array[Index].Indexes.Length) && ((array[Index].Indexes[LigLen - 1] + 1) == array[Index].Indexes[LigLen]))
                                {
                                    LigLen++;
                                }
                                if (LigLen != 1)
                                {
                                    int CheckGlyphCount = 0;
                                    short[] CheckClusterMap = new short[((RunCount + LigLen) - 1) + 1];
                                    SharpDX.DirectWrite.ShapingTextProperties[] CheckTextProps = new SharpDX.DirectWrite.ShapingTextProperties[((RunCount + LigLen) - 1) + 1];
                                    short[] CheckGlyphIndices = new short[(maxGlyphCount - 1) + 1];
                                    SharpDX.DirectWrite.ShapingGlyphProperties[] CheckGlyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
                                    analyzer.GetGlyphs(Str.Substring(RunStart, RunCount + LigLen), RunCount + LigLen, fontFace, false, IsRTL, scriptAnalysis, null, null, new SharpDX.DirectWrite.FontFeature[][] { featureArray }, new int[] { RunCount + LigLen }, maxGlyphCount, CheckClusterMap, CheckTextProps, CheckGlyphIndices, CheckGlyphProps, out CheckGlyphCount);
                                    if ((CheckGlyphCount != LigLen) & (CheckGlyphCount != (LigLen - (((glyphProps[RunRes].Justification != SharpDX.DirectWrite.ScriptJustify.Blank) & (glyphProps[RunRes].Justification != SharpDX.DirectWrite.ScriptJustify.ArabicBlank)) ? 0 : 1))))
                                    {
                                        LigLen = 1;
                                    }
                                }
                            }
                            if ((!glyphProps[ResCount].IsDiacritic | !glyphProps[ResCount].IsZeroWidthSpace) | !glyphProps[ResCount].IsClusterStart)
                            {
                                CharPosInfo info;
                                if (((LigLen == 1) && System.Text.RegularExpressions.Regex.Match(Str[RunStart + RunCount].ToString(), @"[\p{IsArabic}\p{IsArabicPresentationForms-A}\p{IsArabicPresentationForms-B}]").Success) & (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(Str[RunStart + RunCount]) == System.Globalization.UnicodeCategory.DecimalDigitNumber))
                                {
                                    SharpDX.DirectWrite.GlyphMetrics[] _Mets = fontFace.GetDesignGlyphMetrics(glyphIndices, false);
                                    info = new CharPosInfo
                                    {
                                        Index = RunStart + RunCount,
                                        Length = (Index == -1) ? 1 : LigLen,
                                        PriorWidth = PriorWidth,
                                        Width = 2f * ((_Mets[ResCount].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)),
                                        X = (glyphOffsets[ResCount].AdvanceOffset - glyphAdvances[RunRes]) - (((_Mets[ResCount].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) / 4f),
                                        Y = glyphOffsets[ResCount].AscenderOffset,
                                        Height = (((_Mets[ResCount].AdvanceHeight + _Mets[ResCount].BottomSideBearing) - _Mets[ResCount].TopSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)
                                    };
                                    list.Add(info);
                                }
                                else
                                {
                                    SharpDX.DirectWrite.GlyphMetrics[] _Mets = fontFace.GetDesignGlyphMetrics(glyphIndices, false);
                                    info = new CharPosInfo
                                    {
                                        Index = RunStart + RunCount,
                                        Length = (Index == -1) ? 1 : LigLen,
                                        PriorWidth = PriorWidth - ((((glyphProps[RunRes].Justification == SharpDX.DirectWrite.ScriptJustify.ArabicKashida) & (RunCount == 1)) & ((((CharCount == (clusterMap.Length - 1)) ? actualGlyphCount : clusterMap[CharCount + 1]) - clusterMap[CharCount]) == (CharCount - RunStart))) ? glyphAdvances[RunRes] : 0f),
                                        Width = glyphAdvances[RunRes] + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((_Mets[RunRes].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f),
                                        X = glyphOffsets[ResCount].AdvanceOffset,
                                        Y = glyphOffsets[ResCount].AscenderOffset + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((((_Mets[RunRes].AdvanceHeight - _Mets[RunRes].TopSideBearing) - _Mets[RunRes].VerticalOriginY) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f)
                                    };
                                    list.Add(info);
                                    if (((glyphProps[RunRes].Justification == SharpDX.DirectWrite.ScriptJustify.ArabicKashida) & (RunCount == 1)) & ((((CharCount == (clusterMap.Length - 1)) ? actualGlyphCount : clusterMap[CharCount + 1]) - clusterMap[CharCount]) == (CharCount - RunStart)))
                                    {
                                        info = new CharPosInfo
                                        {
                                            Index = (RunStart + RunCount) + 1,
                                            Length = (Index == -1) ? 1 : LigLen,
                                            PriorWidth = PriorWidth,
                                            Width = glyphAdvances[RunRes] + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((_Mets[RunRes].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f),
                                            X = glyphOffsets[ResCount].AdvanceOffset,
                                            Y = glyphOffsets[RunRes].AscenderOffset + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((((_Mets[RunRes].AdvanceHeight - _Mets[RunRes].TopSideBearing) - _Mets[RunRes].VerticalOriginY) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f)
                                        };
                                        list.Add(info);
                                    }
                                }
                            }
                            else
                            {
                                PriorWidth -= glyphOffsets[ResCount].AdvanceOffset;
                            }
                        }
                        if ((CharCount == (clusterMap.Length - 1)) || (clusterMap[CharCount] != clusterMap[CharCount + 1]))
                        {
                            PriorWidth += glyphAdvances[ResCount];
                            int Index = Array.FindIndex<XMLRender.ArabicData.LigatureInfo>(array, (lig) => lig.Indexes[0] == RunStart);
                            if ((Index == -1) || ((((glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.Blank) & (glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.ArabicBlank)) | (Array.IndexOf<int>(array[Index].Indexes, RunStart) == -1)) & ((RunStart + RunCount) != (Str.Length - 1))))
                            {
                                RunCount++;
                            }
                            if ((Index != -1) && (((glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.Blank) & (glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.ArabicBlank)) | (Array.IndexOf<int>(array[Index].Indexes, RunStart) == -1)))
                            {
                                while ((Array.IndexOf<int>(array[Index].Indexes, RunStart + RunCount) != -1) & ((RunStart + RunCount) != (Str.Length - 1)))
                                {
                                    RunCount++;
                                }
                            }
                            if ((clusterMap[CharCount] != ResCount) & !(glyphAdvances[ResCount] == 0f))
                            {
                                RunStart = CharCount;
                                RunCount = 0;
                                RunRes = ResCount;
                            }
                        }
                    }
                    if ((CharCount != (clusterMap.Length - 1)) && (clusterMap[CharCount] != clusterMap[CharCount + 1]))
                    {
                        RunStart = CharCount + 1;
                        if (!(glyphAdvances[clusterMap[CharCount + 1]] == 0f) | (glyphProps[clusterMap[CharCount + 1]].IsClusterStart & glyphProps[clusterMap[CharCount + 1]].IsDiacritic))
                        {
                            RunRes = clusterMap[CharCount + 1];
                        }
                    }
                }
            }
            float Width = 0f;
            float Top = 0f;
            float Bottom = 0f;
            SharpDX.DirectWrite.GlyphMetrics[] designGlyphMetrics = fontFace.GetDesignGlyphMetrics(glyphIndices, false);
            float Left = IsRTL ? 0f : (glyphOffsets[0].AdvanceOffset - ((designGlyphMetrics[0].LeftSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
            float Right = IsRTL ? (glyphOffsets[0].AdvanceOffset - ((designGlyphMetrics[0].RightSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm))) : 0f;
            for (int i = 0; i <= designGlyphMetrics.Length - 1; i++)
            {
                Left = IsRTL ? Math.Max(Left, (glyphOffsets[i].AdvanceOffset + Width) - ((Math.Max(0, designGlyphMetrics[i].LeftSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm))) : Math.Min(Left, (glyphOffsets[i].AdvanceOffset + Width) - ((designGlyphMetrics[i].LeftSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
                if (!(glyphAdvances[i] == 0f))
                {
                    Width += (IsRTL ? ((float)(-1)) : ((float)1)) * ((designGlyphMetrics[i].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm));
                }
                Right = IsRTL ? Math.Min(Right, (glyphOffsets[i].AdvanceOffset + Width) - ((designGlyphMetrics[i].RightSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm))) : Math.Max(Right, (glyphOffsets[i].AdvanceOffset + Width) - ((Math.Min(0, designGlyphMetrics[i].RightSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
                Top = Math.Max(Top, glyphOffsets[i].AscenderOffset + (((designGlyphMetrics[i].VerticalOriginY - designGlyphMetrics[i].TopSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
                Bottom = Math.Min(Bottom, glyphOffsets[i].AscenderOffset + ((((designGlyphMetrics[i].VerticalOriginY - designGlyphMetrics[i].AdvanceHeight) + designGlyphMetrics[i].BottomSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
            }
            Pos = list.ToArray();
            Size Size = new Size(IsRTL ? (Left - Right) : (Right - Left), (Top - Bottom) + ((fontFace.Metrics.LineGap * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
            BaseLine = Top;
            analysisSource.Shadow.Dispose();
            analysisSink.Shadow.Dispose();
            analysisSource.Dispose();
            analysisSource._Factory = null;
            analysisSink.Dispose();
            fontFace.Dispose();
            font.Dispose();
            analyzer.Dispose();
            return Size;
        }
    }
#else
    public static class TextShaping
    {
        public static void Cleanup(int AllNormArb)
        {
        }
        public static double CalculateWidth(string text, bool IsArabic, float maxWidth, float maxHeight)
        {
            double width = 0; //(text, maxWidth, MyUIChanger.FontSize, MyUIChanger.FontFamily);
            return width;
        }
        public static short[] GetWordDiacriticClusters(string Str, string useFont, float fontSize, bool IsRTL)
        {
            return new short[] {};
        }
        public static List<string> GetFontList()
        {
            List<string> fontList = new List<string>();
            return fontList;
        }
        public static string GetAppLanguage()
        {
            return string.Empty;
        }
        public static void SetAppLanguage(string Lang)
        {
        }
        public static string GetSelectedAppLanguage()
        {
            return string.Empty;
        }
        public static List<string> GetAppLanguageList()
        {
            return new List<string>();
        }
    }
#endif
    public class AppSettings : INotifyPropertyChanged
    {
        public static XMLRender.PortableMethods _PortableMethods;
        public static XMLRender.ArabicData ArbData;
        public static IslamMetadata.Arabic Arb;
        public static IslamMetadata.CachedData ChData;
        public static IslamMetadata.TanzilReader TR;
#if WINDOWS_APP || WINDOWS_PHONE_APP
        public static bool ContainsKey(string Key) { return Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey(Key); }
        public static object GetValue(string Key) { return Windows.Storage.ApplicationData.Current.LocalSettings.Values[Key]; }
        public static void SetValue(string Key, object Value) { Windows.Storage.ApplicationData.Current.LocalSettings.Values[Key] = Value; }
#elif WINDOWS_PHONE
        public static bool ContainsKey(string Key) { return false; }
        public static object GetValue(string Key) { return null; }
        public static void SetValue(string Key, object Value) { }
#else
        public static bool ContainsKey(string Key) { return false; }
        public static object GetValue(string Key) { return null; }
        public static void SetValue(string Key, object Value) { }
#endif
        public AppSettings() { }
        public async static System.Threading.Tasks.Task InitDefaultSettings()
        {
#if WINDOWS_APP || WINDOWS_PHONE_APP
            _PortableMethods = new XMLRender.PortableMethods(new WindowsRTFileIO(), new WindowsRTSettings());
#elif WINDOWS_PHONE
            _PortableMethods = new XMLRender.PortableMethods(new WindowsRTXamFileIO(), new WindowsRTXamSettings());
#else
            _PortableMethods = new XMLRender.PortableMethods(new AndroidiOSFileIO(), new AndroidiOSSettings());
#endif
            await _PortableMethods.Init();
            ArbData = new XMLRender.ArabicData(_PortableMethods);
            await ArbData.Init();
            Arb = new IslamMetadata.Arabic(_PortableMethods, ArbData);
            ChData = new IslamMetadata.CachedData(_PortableMethods, ArbData, Arb);
            await ChData.Init(false);
            await Arb.Init(ChData, false);
            TR = new IslamMetadata.TanzilReader(_PortableMethods, Arb, ArbData, ChData);
            await TR.Init();
            if (!ContainsKey("CurrentFont"))
            {
                strSelectedFont = "Times New Roman";
            }
            if (!ContainsKey("OtherCurrentFont"))
            {
                strOtherSelectedFont = "Arial";
            }
            if (!ContainsKey("FontSize"))
            {
                dFontSize = 30.0;
            }
            if (!ContainsKey("OtherFontSize"))
            {
                dOtherFontSize = 20.0;
            }
            if (!ContainsKey("UseColoring"))
            {
                bUseColoring = true;
            }
            if (!ContainsKey("ShowTranslation"))
            {
                bShowTranslation = true;
            }
            if (!ContainsKey("ShowTransliteration"))
            {
                bShowTransliteration = true;
            }
            if (!ContainsKey("ShowW4W"))
            {
                bShowW4W = true;
            }
            if (!ContainsKey("CurrentTranslation"))
            {
                iSelectedTranslation = AppSettings.TR.GetTranslationIndex(String.Empty);
            }
            if (!ContainsKey("CurrentReciter"))
            {
                iSelectedReciter = IslamMetadata.AudioRecitation.GetReciterIndex(String.Empty, AppSettings.ChData.IslamData.ReciterList);
            }
            if (!ContainsKey("Bookmarks"))
            {
                Bookmarks = new int[][] { };
            }
            if (!ContainsKey("DefaultStartTab"))
            {
                iDefaultStartTab = 1;
            }
            if (!ContainsKey("AutomaticAdvanceVerse"))
            {
                bAutomaticAdvanceVerse = true;
            }
            if (!ContainsKey("DelayVerseLengthBeforeAdvancing"))
            {
                bDelayVerseLengthBeforeAdvancing = false;
            }
            if (!ContainsKey("AdditionalVerseAdvanceDelay"))
            {
                iAdditionalVerseAdvanceDelay = 0;
            }
            if (!ContainsKey("LoopingMode"))
            {
                LoopingMode = AppSettings.ChData.IslamData.LoopingModeList.DefaultLoopingMode;
            }
        }
        public static string LoopingMode { get { return (string)GetValue("LoopingMode"); } set { SetValue("LoopingMode", value); } }
        public List<ComboPair> LoopingTypes
        {
            get
            {
                return new List<ComboPair>(AppSettings.ChData.IslamData.LoopingModeList.LoopingModes.Select((Mode) => { return new ComboPair() { KeyString = Mode.Name, ValueString = AppSettings._PortableMethods.LoadResourceString("IslamInfo_" + Mode.Name) }; }));
            }
        }
        public static bool bAutomaticAdvanceVerse { get { return (bool)GetValue("AutomaticAdvanceVerse"); } set { SetValue("AutomaticAdvanceVerse", value); } }
        public bool AutomaticAdvanceVerse { get { return bAutomaticAdvanceVerse; } set { bAutomaticAdvanceVerse = value; } }
        public static bool bDelayVerseLengthBeforeAdvancing { get { return (bool)GetValue("DelayVerseLengthBeforeAdvancing"); } set { SetValue("DelayVerseLengthBeforeAdvancing", value); } }
        public bool DelayVerseLengthBeforeAdvancing { get { return bDelayVerseLengthBeforeAdvancing; } set { bDelayVerseLengthBeforeAdvancing = value; } }
        public static int iAdditionalVerseAdvanceDelay { get { return (int)GetValue("AdditionalVerseAdvanceDelay"); } set { SetValue("AdditionalVerseAdvanceDelay", value); } }
        public string AdditionalVerseAdvanceDelay { get { return iAdditionalVerseAdvanceDelay.ToString(); } set { if (value != null) { iAdditionalVerseAdvanceDelay = int.Parse(value); } } }
        public static int iDefaultStartTab { get { return (int)GetValue("DefaultStartTab"); } set { SetValue("DefaultStartTab", value); } }
        public static int[][] Bookmarks { get { return !((string)GetValue("Bookmarks")).Contains(',') ? new int[][] { } : System.Linq.Enumerable.Select(((string)GetValue("Bookmarks")).Split('|'), (Bookmark) => System.Linq.Enumerable.Select(Bookmark.Split(','), (Str) => int.Parse(Str)).ToArray()).ToArray(); } set { SetValue("Bookmarks", string.Join("|", System.Linq.Enumerable.Select(value, (Bookmark) => string.Join(",", System.Linq.Enumerable.Select(Bookmark, (Mark) => Mark.ToString()))))); } }
#if __ANDROID__ || __IOS__
        public static string strSelectedFont { get { return Plugin.Settings.CrossSettings.Current.GetValueOrDefault("CurrentFont", "Times New Roman"); } set { Plugin.Settings.CrossSettings.Current.AddOrUpdateValue("CurrentFont", value); } }
#else
        public static string strSelectedFont { get { return (string)GetValue("CurrentFont"); } set { SetValue("CurrentFont", value); } }
#endif
        public static string strOtherSelectedFont { get { return (string)GetValue("OtherCurrentFont"); } set { SetValue("OtherCurrentFont", value); } }
        public string SelectedFont { get { return strSelectedFont; } set { strSelectedFont = value; } }
        public string OtherSelectedFont { get { return strOtherSelectedFont; } set { strOtherSelectedFont = value; } }

        public static double dFontSize { get { return (double)GetValue("FontSize"); } set { SetValue("FontSize", value); } }
        public static double dOtherFontSize { get { return (double)GetValue("OtherFontSize"); } set { SetValue("OtherFontSize", value); } }
        public string FontSize { get { return dFontSize.ToString(); } set { double fontSize; if (double.TryParse(value, out fontSize)) { dFontSize = fontSize; } } }
        public string OtherFontSize { get { return dOtherFontSize.ToString(); } set { double fontSize; if (double.TryParse(value, out fontSize)) { dOtherFontSize = fontSize; } } }
        public List<string> GetFontList()
        {
            return TextShaping.GetFontList(); // new List<string> { "Times New Roman", "Traditional Arabic", "Arabic Typesetting", "Sakkal Majalla", "Microsoft Uighur", "Arial", "Global User Interface" };
        }
        public List<string> Fonts
        {
            get
            {
                return GetFontList();
            }
        }
        public List<string> OtherFonts
        {
            get
            {
                return GetFontList();
            }
        }
        public static int iSelectedReciter { get { return (int)GetValue("CurrentReciter"); } set { SetValue("CurrentReciter", value); } }
        public class ComboPair
        {
            public string KeyString { get; set; }
            public string ValueString { get; set; }
        }
        public ComboPair SelectedReciter { get { return ReciterList.First((Item) => Item.KeyString == AppSettings.ChData.IslamData.ReciterList.Reciters[iSelectedReciter].Name); } set { if (value != null) { iSelectedReciter = Array.FindIndex(AppSettings.ChData.IslamData.ReciterList.Reciters, (Reciter) => Reciter.Name == value.KeyString); } } }
        public List<ComboPair> _ReciterList;
        public List<ComboPair> ReciterList
        {
            get
            {
                if (_ReciterList == null) _ReciterList = new List<ComboPair>(AppSettings.ChData.IslamData.ReciterList.Reciters.Select((Reciter) => { return new ComboPair() { KeyString = Reciter.Name, ValueString = Reciter.Reciter + (Reciter.BitRate == 0 ? string.Empty : (" [" + Reciter.BitRate.ToString() + "kbps]")) }; }));
                return _ReciterList;
            }
        }

        public static int iSelectedTranslation { get { return (int)GetValue("CurrentTranslation"); } set { SetValue("CurrentTranslation", value); } }
        public string SelectedTranslation { get { return AppSettings.ChData.IslamData.Translations.TranslationList[iSelectedTranslation].Name; } set { if (value != null) { iSelectedTranslation = Array.FindIndex(AppSettings.ChData.IslamData.Translations.TranslationList, (Translation) => Translation.Name == value); } } }
        public List<string> TranslationList
        {
            get
            {
                return new List<string>(AppSettings.ChData.IslamData.Translations.TranslationList.Where((Translation) => { return Translation.FileName.Substring(0, Translation.FileName.IndexOf('.')) == System.Globalization.CultureInfo.CurrentCulture.Name || Translation.FileName.Substring(0, Translation.FileName.IndexOf('.')) == System.Globalization.CultureInfo.CurrentCulture.Parent.Name; }).Select((Translation) => Translation.Name));
            }
        }
        public static bool bUseColoring { get { return (bool)GetValue("UseColoring"); } set { SetValue("UseColoring", value); } }
        public static bool bShowTranslation { get { return (bool)GetValue("ShowTranslation"); } set { SetValue("ShowTranslation", value); } }
        public static bool bShowTransliteration { get { return (bool)GetValue("ShowTransliteration"); } set { SetValue("ShowTransliteration", value); } }
        public static bool bShowW4W { get { return (bool)GetValue("ShowW4W"); } set { SetValue("ShowW4W", value); } }
        public bool UseColoring { get { return bUseColoring; } set { bUseColoring = value; } }
        public bool ShowTranslation { get { return bShowTranslation; } set { bShowTranslation = value; } }
        public bool ShowTransliteration { get { return bShowTransliteration; } set { bShowTransliteration = value; } }
        public bool ShowW4W { get { return bShowW4W; } set { bShowW4W = value; } }
        public static string strAppLanguage
        {
            get { return TextShaping.GetAppLanguage(); }
            set
            {
                TextShaping.SetAppLanguage(value);
            }
        }
        public string AppLanguage { get { return TextShaping.GetSelectedAppLanguage(); } set { strAppLanguage = value.Substring(value.LastIndexOf("(")).Trim('(', ')'); PropertyChanged(this, new PropertyChangedEventArgs("TranslationList")); iSelectedTranslation = AppSettings.TR.GetTranslationIndex(String.Empty); PropertyChanged(this, new PropertyChangedEventArgs("SelectedTranslation")); } }

        public List<string> AppLanguageList
        {
            get
            {
                return TextShaping.GetAppLanguageList();
            }
        }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
    public class MyTabViewModel : INotifyPropertyChanged
    {
        public MyTabViewModel()
        {
        }
        private List<MyTabItem> _Items;
        public IEnumerable<MyTabItem> Items
        {
            get
            {
                return _Items;
            }
            set
            {
                _Items = value.ToList();
                int iDefault = AppSettings.iDefaultStartTab;
                PropertyChanged(this, new PropertyChangedEventArgs("Items"));
                SelectedItem = _Items.ElementAt(iDefault);
            }
        }

        private List<MyListItem> _ListItems;
        public IEnumerable<MyListItem> ListItems
        {
            get
            {
                return _ListItems;
            }
            set
            {
                _ListItems = value.ToList();
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("ListItems"));
            }
        }

        public MyListItem ListSelectedItem
        {
            get { return _selectedItem == null ? null : _selectedItem.SelectedItem; }
            set
            {
                _selectedItem.SelectedItem = value;
                PropertyChanged(this, new PropertyChangedEventArgs("ListSelectedItem"));
            }
        }

        private MyTabItem _selectedItem;

        public MyTabItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                AppSettings.iDefaultStartTab = _selectedItem.Index;
                PropertyChanged(this, new PropertyChangedEventArgs("SelectedItem"));
                ListItems = _selectedItem.Items;
                ListSelectedItem = ListItems.Count() == 0 ? null : ListItems.First();
            }
        }

#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }

    public class MyTabItem
    {
        public bool IsBookmarks { get; set; }
        public string Title { get; set; }
        public int Index { get; set; }
        private IEnumerable<MyListItem> _Items;
        public IEnumerable<MyListItem> Items
        {
            get
            {
                if (_Items == null)
                {
                    if (IsBookmarks)
                    {
                        _Items = System.Linq.Enumerable.Select(AppSettings.Bookmarks, (Bookmark, Idx) => new MyListItem { TextItems = new List<string>() { AppSettings.TR.GetSelectionName(Bookmark[0], Bookmark[1], XMLRender.ArabicData.TranslitScheme.RuleBased, String.Empty) + " " + Bookmark[2].ToString() + ":" + Bookmark[3].ToString() }, Index = Idx });
                    }
                    else
                    {
                        _Items = System.Linq.Enumerable.Select(AppSettings.TR.GetSelectionNames((Index - 1).ToString(), XMLRender.ArabicData.TranslitScheme.RuleBased, String.Empty), (Arr, Idx) => new MyListItem { TextItems = new List<string>(((string)(Arr.Cast<object>()).First()).Split('(', ')')), Index = (int)(Arr.Cast<object>()).Last() });
                    }
                }
                return _Items;
            }
        }
        public void RefreshItems() { _Items = null; }

        private MyListItem _selectedItem;

        public MyListItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                if (_selectedItem != null) _selectedItem.IsSelected = false;
                _selectedItem = value;
                if (_selectedItem != null) _selectedItem.IsSelected = true;
            }
        }
    }

    public class MyListItem : INotifyPropertyChanged
    {
        public List<string> TextItems { get; set; }
        public int Index { get; set; }
        private bool _IsSelected;
        public bool IsSelected { get { return _IsSelected; } set { _IsSelected = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsSelected")); } }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
    public class MyUIChanger : INotifyPropertyChanged
    {
        public MyUIChanger()
        {
        }
        public double FontSize
        {
            get
            {
                return AppSettings.dFontSize;
            }
            set
            {
                AppSettings.dFontSize = value;
                TextShaping.Cleanup(2);
                PropertyChanged(this, new PropertyChangedEventArgs("FontSize"));
            }
        }
        public string FontFamily
        {
            get
            {
                return AppSettings.strSelectedFont;
            }
        }
        public double OtherFontSize
        {
            get
            {
                return AppSettings.dOtherFontSize;
            }
            set
            {
                AppSettings.dOtherFontSize = value;
                TextShaping.Cleanup(1);
                PropertyChanged(this, new PropertyChangedEventArgs("OtherFontSize"));
            }
        }
        public string OtherFontFamily
        {
            get
            {
                return AppSettings.strOtherSelectedFont;
            }
        }
        private double _MaxWidth;
        public double MaxWidth
        {
            get
            {
                return _MaxWidth;
            }
            set
            {
                _MaxWidth = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth"));
            }
        }
        private double _MaxHeight;
        public double MaxHeight
        {
            get
            {
                return _MaxHeight;
            }
            set
            {
                _MaxHeight = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxHeight"));
            }
        }

#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
    public class MyRenderModel : INotifyPropertyChanged
    {
        public MyRenderModel(List<MyRenderItem> NewRenderItems)
        {
            RenderItems = NewRenderItems;
            MaxWidth = CalculateWidth();
        }
        private double _MaxWidth;
        public double MaxWidth { get { return _MaxWidth; } set { _MaxWidth = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth")); } }
        private double CalculateWidth()
        {
            return RenderItems.Select((Item) => Item.MaxWidth).Sum();
        }
        private List<MyRenderItem> _RenderItems;
        public List<MyRenderItem> RenderItems { get { return _RenderItems; } set { _RenderItems = value.ToList(); if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("RenderItems")); } }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
    public class MyRenderItem : INotifyPropertyChanged
    {
        public MyRenderItem(XMLRender.RenderArray.RenderItem RendItem)
        {
            if (RendItem.TextItems.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eReference) { Chapter = ((int[])RendItem.TextItems.First().Text)[0]; Verse = ((int[])RendItem.TextItems.First().Text)[1]; Word = ((int[])RendItem.TextItems.First().Text).Count() == 2 ? -1 : ((int[])RendItem.TextItems.First().Text)[2]; } else { Chapter = -1; Verse = -1; Word = -1; }
            Items = System.Linq.Enumerable.Select(RendItem.TextItems.GroupBy((MainItems) => (MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eLTR || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eTransliteration) ? (object)MainItems.DisplayClass : (object)MainItems), (Arr) => (Arr.First().Text.GetType() == typeof(List<XMLRender.RenderArray.RenderItem>)) ? (object)new VirtualizingWrapPanelAdapter() { RenderModels = new List<MyRenderModel>() { new MyRenderModel(System.Linq.Enumerable.Select((List<XMLRender.RenderArray.RenderItem>)Arr.First().Text, (ArrRend) => new MyRenderItem((XMLRender.RenderArray.RenderItem)ArrRend)).ToList()) } } : ((Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eContinueStop) ? (object)new MyChildRenderStopContinue((bool)((object[])Arr.First().Text)[0], (List<IslamMetadata.Arabic.RuleMetadata>)((object[])Arr.First().Text)[1]) : (Arr.First().Text.GetType() == typeof(string) ? (object)new MyChildRenderItem(System.Linq.Enumerable.Select(Arr, (ArrItem) => new MyChildRenderBlockItem() { ItemText = (string)ArrItem.Text, Clr = ArrItem.Clr }).ToList(), Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic, Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL) : null))).Where(Arr => Arr != null).ToList();
            MaxWidth = CalculateWidth();
        }
        public double MaxWidth { get; set; }
        private double CalculateWidth()
        {
            if (Items.Count() == 0) return 0.0;
            return _Items.Select((Item) => Item.GetType() == typeof(MyChildRenderItem) ? ((MyChildRenderItem)Item).MaxWidth : (Item.GetType() == typeof(MyChildRenderStopContinue) ? ((MyChildRenderStopContinue)Item).MaxWidth : ((VirtualizingWrapPanelAdapter)Item).RenderModels.Select((It) => It.MaxWidth).Max())).Max();
        }
        public void RegroupRenderModels(double maxWidth)
        {
            _Items.FirstOrDefault((Item) => { if (Item.GetType() != typeof(MyChildRenderItem) && Item.GetType() != typeof(MyChildRenderStopContinue)) { ((VirtualizingWrapPanelAdapter)Item).RegroupRenderModels(maxWidth); } return false; });
        }
        private List<object> _Items;
        public List<object> Items { get { return _Items; } set { _Items = value.ToList(); if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Items")); } }
        private bool _IsSelected;
        public bool IsSelected
        {
            get { return _IsSelected; }
            set
            {
                _IsSelected = value;
                for (int count = 0; count < _Items.Count; count++) { if (_Items[count].GetType() != typeof(MyChildRenderItem) && _Items[count].GetType() != typeof(MyChildRenderStopContinue)) { for (int subc = 0; subc < ((VirtualizingWrapPanelAdapter)_Items[count]).RenderModels.Count; subc++) { for (int itc = 0; itc < ((VirtualizingWrapPanelAdapter)_Items[count]).RenderModels[subc].RenderItems.Count; itc++) { ((VirtualizingWrapPanelAdapter)_Items[count]).RenderModels[subc].RenderItems[itc].IsSelected = value; } } } }
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsSelected"));
            }
        }
        public int Chapter;
        public int Verse;
        public int Word;
        public string VerseText { get { return Chapter.ToString() + ":" + Verse.ToString(); } }
        public string VerseWordText { get { return Chapter.ToString() + ":" + Verse.ToString() + (Word == -1 ? string.Empty : (":" + Word.ToString())); } }
        public string GetText { get { return String.Join(" ", _Items.Select((Item) => Item.GetType() == typeof(MyChildRenderItem) ? ((MyChildRenderItem)Item).GetText : String.Join(string.Empty, ((VirtualizingWrapPanelAdapter)Item).RenderModels.Select((It) => String.Join(string.Empty, It.RenderItems.Select((RecIt) => RecIt.GetText)))))) + " " + VerseWordText; } }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
    public class VirtualizingWrapPanelAdapter : INotifyPropertyChanged
    {
        public VirtualizingWrapPanelAdapter RenderSource { get { return this; } }
        private List<MyRenderModel> _RenderModels;
        public List<MyRenderModel> RenderModels
        {
            get
            {
                if (_RenderModels == null) { _RenderModels = new List<MyRenderModel>(); }
                return _RenderModels;
            }
            set
            {
                _RenderModels = value;
                VerseReferences = value.SelectMany((Model) => Model.RenderItems).Where((Item) => Item.Chapter != -1).ToList();
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("RenderModels"));
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("RenderSource"));
            }
        }
        public void RegroupRenderModels(double maxWidth)
        {
            if (_RenderModels != null) { RenderModels = GroupRenderModels(_RenderModels.SelectMany((Item) => Item.RenderItems).ToList(), maxWidth); }
        }
        public static List<MyRenderModel> GroupRenderModels(List<MyRenderItem> value, double maxWidth)
        {
            double width = 0.0;
            int groupIndex = 0;
            List<int[]> GroupIndexes = new List<int[]>();
            value.FirstOrDefault((Item) => {
                Item.RegroupRenderModels(maxWidth); double itemWidth = Math.Min(Item.MaxWidth, maxWidth); if (width + itemWidth > maxWidth) { width = itemWidth; groupIndex++; } else { width += itemWidth; }
                GroupIndexes.Add(new int[] { groupIndex, GroupIndexes.Count }); return false;
            });
            return GroupIndexes.GroupBy((Item) => Item[0], (Item) => value.ElementAt(Item[1])).Select((Item) => new MyRenderModel(Item.ToList())).ToList();
        }
        private List<MyRenderItem> _VerseReferences;
        public List<MyRenderItem> VerseReferences { get { return _VerseReferences; } set { _VerseReferences = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("VerseReferences")); } }
        private MyRenderItem _CurrentVerse;
        public MyRenderItem CurrentVerse { get { return _CurrentVerse; } set { if (_CurrentVerse != null) _CurrentVerse.IsSelected = false; _CurrentVerse = value; if (_CurrentVerse != null) _CurrentVerse.IsSelected = true; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("CurrentVerse")); } }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
    public class MyChildRenderStopContinue : INotifyPropertyChanged
    {
        public MyChildRenderStopContinue(bool NewIsStop, List<IslamMetadata.Arabic.RuleMetadata> NewRules) { IsStop = NewIsStop; _Rules = NewRules; MaxWidth = CalculateWidth(); }
        private double _MaxWidth;
        public double MaxWidth { get { return _MaxWidth; } set { _MaxWidth = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth")); } }
        public double CalculateWidth()
        {
            //5 margin on both sides
            return 5 + 5 + 1 + 1 + TextShaping.CalculateWidth(IsStop ? "\u2B59" : "\u2B45", false, (float)float.MaxValue, float.MaxValue);
        }
        private List<IslamMetadata.Arabic.RuleMetadata> _Rules;
        public List<IslamMetadata.Arabic.RuleMetadata> MetaRules { get { return _Rules; } }
        private bool _IsStop;
        public bool IsStop { get { return _IsStop; } set { _IsStop = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsStop")); } }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
    public class MyChildRenderItem : INotifyPropertyChanged
    {
        private double _MaxWidth;
        public double MaxWidth { get { return _MaxWidth; } set { _MaxWidth = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth")); } }

        public double CalculateWidth()
        {
            //5 margin on both sides
            return 5 + 5 + 1 + 1 + TextShaping.CalculateWidth(string.Join(String.Empty, ItemRuns.Select((Item) => Item.ItemText)), IsArabic, (float)float.MaxValue, float.MaxValue);
        }
        public string GetText { get { return String.Join(string.Empty, _ItemRuns.Select((Item) => Item.ItemText)); } }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion

        public MyChildRenderItem(List<MyChildRenderBlockItem> NewItemRuns, bool NewIsArabic, bool NewIsRTL)
        {
            IsRTL = NewIsRTL;
            IsArabic = NewIsArabic; //must be set before ItemRuns
            ItemRuns = NewItemRuns;
            MaxWidth = CalculateWidth();
        }
        List<MyChildRenderBlockItem> _ItemRuns;
        public List<MyChildRenderBlockItem> ItemRuns
        {
            get { return _ItemRuns; }
            set
            {
                if (IsArabic)
                {
                    char[] Forms = AppSettings.ArbData.GetPresentationForms;
                    //XMLRender.ArabicData.LigatureInfo[] ligs = XMLRender.ArabicData.GetLigatures(String.Join(String.Empty, System.Linq.Enumerable.Select(value, (Run) => Run.ItemText)), false, Forms);
                    short[] chpos = TextShaping.GetWordDiacriticClusters(String.Join(String.Empty, System.Linq.Enumerable.Select(value, (Run) => Run.ItemText)), AppSettings.strSelectedFont, (float)AppSettings.dFontSize, IsArabic);
                    int pos = value[0].ItemText.Length;
                    int count = 1;
                    while (count < value.Count)
                    {
                        if (value[count].ItemText.Length != 0 && pos != 0 && chpos[pos] == chpos[pos - 1])
                        {
                            if (!AppSettings.ArbData.IsDiacritic(value[count].ItemText.First()) && ((value[count - 1].Clr & 0xFFFFFF) != 0x000000)) { value[count - 1].Clr = value[count].Clr; }
                            if (chpos[pos] == chpos[pos + value[count].ItemText.Length - 1])
                            {
                                pos -= value[count - 1].ItemText.Length;
                                value[count - 1].ItemText += value[count].ItemText; value.RemoveAt(count); count--;
                            }
                            else
                            {
                                int subcount = pos;
                                while (subcount < pos + value[count].ItemText.Length && chpos[subcount] == chpos[subcount + 1]) { subcount++; }
                                value[count - 1].ItemText += value[count].ItemText.Substring(0, subcount - pos + 1); value[count].ItemText = value[count].ItemText.Substring(subcount - pos + 1);
                                pos += subcount - pos + 1;
                            }
                        }
                        //for (int subcount = 0; subcount < ligs.Length; subcount++) {
                        //    //for indexes which are before, and after, find the maximum index in this group which is after
                        //    if (pos > ligs[subcount].Indexes[0] && pos <= ligs[subcount].Indexes[ligs[subcount].Indexes.Length - 1]) {
                        //        int ligcount;
                        //        for (ligcount = ligs[subcount].Indexes.Length - 1; ligcount >= 0; ligcount--) {
                        //            if (pos + value[count].ItemText.Length <= ligs[subcount].Indexes[ligcount]) {
                        //                if (ligs[subcount].Indexes[ligcount] - pos == value[count].ItemText.Length) {
                        //                    pos -= value[count - 1].ItemText.Length;
                        //                    if (!value[count].ItemText.All((ch) => XMLRender.ArabicData.IsDiacritic(ch))) { value[count - 1].Clr = value[count].Clr; }
                        //                    value[count - 1].ItemText += value[count].ItemText; value.RemoveAt(count); count--;
                        //                } else {
                        //                    value[count - 1].ItemText += value[count].ItemText.Substring(0, ligs[subcount].Indexes[ligcount] - pos); value[count].ItemText = value[count].ItemText.Substring(ligs[subcount].Indexes[ligcount] - pos);
                        //                    pos += ligs[subcount].Indexes[ligcount] - pos;
                        //                }
                        //                break;
                        //            }
                        //        }
                        //        if (ligcount != ligs[subcount].Indexes.Length) { break; }
                        //    }
                        //}
                        pos += value[count].ItemText.Length;
                        count++;
                        //int charcount = 0;
                        //while (charcount < value[count].ItemText.Length && XMLRender.ArabicData.IsDiacritic(value[count].ItemText[charcount])) { charcount++; }
                        //if (charcount != 0)
                        //{
                        //    if (charcount == value[count].ItemText.Length)
                        //    {
                        //        value[count - 1].ItemText += value[count].ItemText; value.RemoveAt(count);
                        //    }
                        //    else { value[count - 1].ItemText += value[count].ItemText.Substring(0, charcount); value[count].ItemText = value[count].ItemText.Substring(charcount); count++; }
                        //}
                        //else { count++; }
                    }
                }
                _ItemRuns = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("ItemRuns"));
            }
        }
        public bool IsArabic { get; set; }
        public bool IsRTL { get; set; }
    }
    public class MyChildRenderBlockItem
    {
        public int Clr;
        public string ItemText { get; set; }
    }
}
