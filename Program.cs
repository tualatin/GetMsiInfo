using System;
using System.Text.RegularExpressions;


namespace GetMSIVersion
{
  class Program
  {
    static string GetInstallerString (WindowsInstaller.View msview)
    {
      WindowsInstaller.Record record = null;

      msview.Execute (record);
      record = msview.Fetch ();

      return (record.get_StringData (1).ToString ());
    }

    static void ShowAllProperties (WindowsInstaller.View msview)
    {
      WindowsInstaller.Record record = null;

      msview.Execute (record);
      record = msview.Fetch ();

      while (record != null)
      {
        Console.Write (string.Format ("\n{0}\t{1}", record.get_StringData (1), record.get_StringData (2)));
        record = msview.Fetch ();
      }
    }

    static void Main (string[] args)
    {
      try
      {
        string strMSIPackage = args[1];

        Type type = Type.GetTypeFromProgID ("WindowsInstaller.Installer");

        WindowsInstaller.Installer installer = (WindowsInstaller.Installer) Activator.CreateInstance (type);
        WindowsInstaller.Database db = installer.OpenDatabase (@strMSIPackage, 0);
        WindowsInstaller.View dv = null;

        string strNumber = string.Empty;

        if ((args[0].CompareTo ("/v")) == 0)
        {
          dv = db.OpenView ("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductVersion'");
          strNumber = GetInstallerString (dv);
          Console.Write (string.Format ("\n{0}", strNumber));
        }
        else if ((args[0].CompareTo ("/n")) == 0)
        {
          dv = db.OpenView ("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductName'");
          strNumber = GetInstallerString (dv);
          Match match = Regex.Match (strNumber, @"\d+(?=\.\w)(.*)", RegexOptions.IgnoreCase);

          if (match.Success)
            Console.Write (string.Format ("\n{0}", (strNumber.Substring (0, strNumber.Length - match.Value.Length)).Trim ( )));
        }
        else if ((args[0].CompareTo ("/a")) == 0)
        {
          dv = db.OpenView ("SELECT `Property`, `Value` FROM `Property`");
          ShowAllProperties (dv);
        }

        // Environment.SetEnvironmentVariable ( "MSIVERSION", strNumber, EnvironmentVariableTarget.User );
      }
      catch (Exception e)
      {
        Console.WriteLine (e.Message);
      }
    }
  }
}
