using System.ComponentModel;

namespace WpfSample;
public class StudentWithScores(
  string name
) : INotifyPropertyChanged {
  public uint Row { get; set; }

  public string Name { get; } = name;

  public string WW1_1 { get; set; } = "";
  public string WW2_1 { get; set; } = "";
  public string WW3_1 { get; set; } = "";
  public string WW4_1 { get; set; } = "";
  public string WW5_1 { get; set; } = "";
  public string WW6_1 { get; set; } = "";
  public string WW7_1 { get; set; } = "";
  public string WW8_1 { get; set; } = "";
  public string WW9_1 { get; set; } = "";
  public string WW10_1 { get; set; } = "";
  
  public string PT1_1 { get; set; } = "";
  public string PT2_1 { get; set; } = "";
  public string PT3_1 { get; set; } = "";
  public string PT4_1 { get; set; } = "";
  public string PT5_1 { get; set; } = "";
  public string PT6_1 { get; set; } = "";
  public string PT7_1 { get; set; } = "";
  public string PT8_1 { get; set; } = "";
  public string PT9_1 { get; set; } = "";
  public string PT10_1 { get; set; } = "";

  public string WW1_2 { get; set; } = "";
  public string WW2_2 { get; set; } = "";
  public string WW3_2 { get; set; } = "";
  public string WW4_2 { get; set; } = "";
  public string WW5_2 { get; set; } = "";
  public string WW6_2 { get; set; } = "";
  public string WW7_2 { get; set; } = "";
  public string WW8_2 { get; set; } = "";
  public string WW9_2 { get; set; } = "";
  public string WW10_2 { get; set; } = "";
  
  public string PT1_2 { get; set; } = "";
  public string PT2_2 { get; set; } = "";
  public string PT3_2 { get; set; } = "";
  public string PT4_2 { get; set; } = "";
  public string PT5_2 { get; set; } = "";
  public string PT6_2 { get; set; } = "";
  public string PT7_2 { get; set; } = "";
  public string PT8_2 { get; set; } = "";
  public string PT9_2 { get; set; } = "";
  public string PT10_2 { get; set; } = "";

  public string EX_1 { get; set; } = "";
  public string EX_2 { get; set; } = "";

  public string Grade_1 { get; set; } = "";
  public string Grade_2 { get; set; } = "";

  private string transmuted_1 = "";
  public string Transmuted_1 { get => transmuted_1; set { transmuted_1 = value; OnPropertyChanged(nameof(Transmuted_1)); } }
  private string transmuted_2 = "";
  public string Transmuted_2 { get => transmuted_2; set { transmuted_2 = value; OnPropertyChanged(nameof(Transmuted_2)); } }

  public event PropertyChangedEventHandler? PropertyChanged;

  protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}


