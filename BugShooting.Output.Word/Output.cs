using BS.Plugin.V3.Output;

namespace BugShooting.Output.Word
{

  public class Output: IOutput 
  {
    
    public string Name
    {
      get { return "Microsoft Word"; }
    }

    public string Information
    {
      get { return string.Empty; }
    }

  }
}
