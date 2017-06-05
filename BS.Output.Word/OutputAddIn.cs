using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;

namespace BS.Output.Word
{
  public class OutputAddIn: V3.OutputAddIn<Output>
  {

    protected override string Name
    {
      get { return "Microsoft Word"; }
    }

    protected override Image Image64x32
    {
      get  { return Properties.Resources.logo_64x32; }
    }

    protected override Image Image16x16
    {
      get { return Properties.Resources.logo_16x16 ; }
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
      return new Output(Name);
    }

    protected override Output EditOutput(IWin32Window Owner, Output Output)
    {
      return null; 
    }

    protected override OutputValueCollection SerializeOutput(Output Output)
    {

      OutputValueCollection outputValues = new OutputValueCollection();

      outputValues.Add(new OutputValue("Name", Output.Name));

      return outputValues;
      
    }

    protected override Output DeserializeOutput(OutputValueCollection OutputValues)
    {
      return new Output(OutputValues["Name", this.Name].Value);
    }

    protected async override Task<V3.SendResult> Send(Output Output, V3.ImageData ImageData)
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

        return new V3.SendResult(V3.Result.Success);
        
      }
      catch (Exception ex)
      {
        return new V3.SendResult(V3.Result.Failed, ex.Message);
      }
      
    }
      
  }

}