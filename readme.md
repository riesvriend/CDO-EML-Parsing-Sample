# EML Parsing Sample in C# using CDO

This sample project answers ([the question on stackoverflow](http://stackoverflow.com/questions/936422/recommendations-on-parsing-eml-files-in-c-sharp)) on how to parse a raw email file in .eml format using C#. 

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

CDOSYS v.6 is [pre-installed](https://support.microsoft.com/en-us/kb/171440) on all modern versions of Windows versions (2000 to 2016). 

# Changelist
20 Jul 2015 Ries Vriend: Initial commit
