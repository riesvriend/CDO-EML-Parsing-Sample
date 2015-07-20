using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace CDO_EML_Parsing_Sample
{
  public class EmlTester
  {
    /// <summary>
    /// Sample to illustrate the use of CDOSYS.DLL from .net to parse EML files
    /// CDOSYS v.6 is a standard part of all windows versions (2012R2, 2008, 2003, 2000, XP, Vista, 8, 10): https://support.microsoft.com/en-us/kb/171440
    /// </summary>
    public static void Main()
    {
      var binFolderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
      var emlFilePath = Path.Combine(binFolderPath, "sample.eml");
      var message = CdoWrapper.LoadMessage(emlFilePath);
      Debug.WriteLine(message.Subject);
      Debug.WriteLine(message.TextBodyPart.GetString());
      Debug.WriteLine(message.HTMLBodyPart.GetString());
      var attachment = message.Attachments[1];
      Debug.WriteLine(attachment.FileName);
      Debug.WriteLine(attachment.GetString());
    }
  }
}
