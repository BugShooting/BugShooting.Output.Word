using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;
using BS.Plugin.V3.Output;
using BS.Plugin.V3.Common;


namespace BugShooting.Output.Word
{
  public class OutputPlugin: OutputPlugin<Output>
  {

    protected override string Name
    {
      get { return "Microsoft Word"; }
    }

    protected override Image Image64
    {
      get  { return Properties.Resources.logo_64; }
    }

    protected override Image Image16
    {
      get { return Properties.Resources.logo_16 ; }
    }

    protected override bool Editable
    {
      get { return false; }
    }

    protected override string Description
    {
      get { return "Insert screenshots into Word documents."; }
    }
    
    protected override Output CreateOutput(IWin32Window Owner)
    {
      return new Output();
    }

    protected override Output EditOutput(IWin32Window Owner, Output Output)
    {
      return null; 
    }

    protected override OutputValues SerializeOutput(Output Output)
    {
      return new OutputValues();
    }

    protected override Output DeserializeOutput(OutputValues OutputValues)
    {
      return new Output();
    }

    protected async override Task<SendResult> Send(IWin32Window Owner, Output Output, ImageData ImageData)
    {
      try
      {

        Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();

        Microsoft.Office.Interop.Word.Document document = application.Documents.Add();

        foreach (Image image in ImageData.Images)
        {

          string filePath = Path.GetTempFileName();

          image.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

          document.InlineShapes.AddPicture(filePath);
          
          File.Delete(filePath);

        }
       
        application.Visible = true;
        document.Activate();

        document = null;
        application = null;

        return new SendResult(Result.Success);
        
      }
      catch (Exception ex)
      {
        return new SendResult(Result.Failed, ex.Message);
      }
      
    }
      
  }

}