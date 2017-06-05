namespace BS.Output.Word
{

  public class Output: IOutput 
  {
    
    string name;

    public Output(string name)
    {
      this.name = name;
    }
    
    public string Name
    {
      get { return name; }
    }

    public string Information
    {
      get { return string.Empty; }
    }

  }
}
