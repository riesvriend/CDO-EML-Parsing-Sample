# CDO-EML-Parsing-Sample

This sample project answers the question ([on stackoverflow](http://stackoverflow.com/questions/936422/recommendations-on-parsing-eml-files-in-c-sharp)) on how .eml files can be parsed in C#.

The sample shows how to read a raw .eml file, and how to extract the body parts to byte arrays or strings. 

This fragment illustrates the API:

```c#
      var message = CdoWrapper.LoadMessage(emlFilePath);
      Debug.WriteLine(message.Subject);
      Debug.WriteLine(message.TextBodyPart.GetString());
      Debug.WriteLine(message.HTMLBodyPart.GetString());
      var attachment = message.Attachments[1];
      Debug.WriteLine(attachment.FileName);
      Debug.WriteLine(attachment.GetString());
```

CDOSYS v.6 is pre-installed on all modern versions of Windows versions (2012, 2008, 2003, 2000): https://support.microsoft.com/en-us/kb/171440

