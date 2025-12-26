using CSharpFunctionalExtensions;
using Microsoft.Win32;
using ModernWpf.Controls;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace WpfSample {
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window {

    private readonly ContentDialog _errorDialog = new ContentDialog {
      Title = "Error",
      CloseButtonText = "Ok"
    };
    private bool _isDialogOpen = false;
    private readonly WorkbookService _workbookService = new WorkbookService();
        
    public ObservableCollection<string> HighestWW_1 { get; } = new ObservableCollection<string>(new string[10]);
    public ObservableCollection<string> HighestPT_1 { get; } = new ObservableCollection<string>(new string[10]);
    private string _highestExam1 = ""; public string HighestExam_1 { 
      get => _highestExam1; 
      set { 
        if (_highestExam1 != value) { 
          _highestExam1 = value;
          txtExam_1.Text = _highestExam1;
          mColExam_1.Header = $"{HighestExam_1}";
          fColExam_1.Header = $"{HighestExam_1}";
        } 
      } 
    }

    public ObservableCollection<string> HighestWW_2 { get; } = new ObservableCollection<string>(new string[10]);
    public ObservableCollection<string> HighestPT_2 { get; } = new ObservableCollection<string>(new string[10]);
    private string _highestExam2 = ""; public string HighestExam_2 { 
      get => _highestExam2; 
      set { 
        if (_highestExam2 != value) { 
          _highestExam2 = value;
          txtExam_2.Text = _highestExam2;
          mColExam_2.Header = $"{HighestExam_2}";  
          fColExam_2.Header = $"{HighestExam_2}"; 
        } 
      } 
    }

    public ObservableCollection<StudentWithScores> Males { get; } = new();

    public ObservableCollection<StudentWithScores> Females { get; } = new();

    public bool Quarter1 { get; set; } = true;
    public bool ShowGrades { get; set; } = true;

    // WW columns
    private DataGridTextColumn[] mWWSet1Columns = [];
    private DataGridTextColumn[] mWWSet2Columns = [];
    private DataGridTextColumn[] fWWSet1Columns = [];
    private DataGridTextColumn[] fWWSet2Columns = [];

    // PT columns
    private DataGridTextColumn[] mPTSet1Columns = [];
    private DataGridTextColumn[] mPTSet2Columns = [];
    private DataGridTextColumn[] fPTSet1Columns = [];
    private DataGridTextColumn[] fPTSet2Columns = [];

    private ICollectionView maleView; 
    private ICollectionView femaleView;

    public MainWindow() {
      InitializeComponent();
      DataContext = this;
      // this.Closed += (s, e) => _workbookService.Dispose(); 
      tabHighestScores.Visibility = Visibility.Hidden;
      tabStudents.Visibility = Visibility.Hidden;
      tabHighestScores.IsEnabled = false;
      tabStudents.IsEnabled = false;

      mWWSet1Columns = [
        mColWW1_1, mColWW2_1, mColWW3_1, mColWW4_1, mColWW5_1, mColWW6_1, mColWW7_1, mColWW8_1, mColWW9_1, mColWW10_1
      ];
      mPTSet1Columns = [
        mColPT1_1, mColPT2_1, mColPT3_1, mColPT4_1, mColPT5_1, mColPT6_1, mColPT7_1, mColPT8_1, mColPT9_1, mColPT10_1
      ];
      mWWSet2Columns = [
        mColWW1_2, mColWW2_2, mColWW3_2, mColWW4_2, mColWW5_2, mColWW6_2, mColWW7_2, mColWW8_2, mColWW9_2, mColWW10_2
      ];
      mPTSet2Columns = [
        mColPT1_2, mColPT2_2, mColPT3_2, mColPT4_2, mColPT5_2, mColPT6_2, mColPT7_2, mColPT8_2, mColPT9_2, mColPT10_2
      ];
      fWWSet1Columns = [
        fColWW1_1, fColWW2_1, fColWW3_1, fColWW4_1, fColWW5_1, fColWW6_1, fColWW7_1, fColWW8_1, fColWW9_1, fColWW10_1
      ];
      fPTSet1Columns = [
        fColPT1_1, fColPT2_1, fColPT3_1, fColPT4_1, fColPT5_1, fColPT6_1, fColPT7_1, fColPT8_1, fColPT9_1, fColPT10_1
      ];
      fWWSet2Columns = [
        fColWW1_2, fColWW2_2, fColWW3_2, fColWW4_2, fColWW5_2, fColWW6_2, fColWW7_2, fColWW8_2, fColWW9_2, fColWW10_2
      ];
      fPTSet2Columns = [
        fColPT1_2, fColPT2_2, fColPT3_2, fColPT4_2, fColPT5_2, fColPT6_2, fColPT7_2, fColPT8_2, fColPT9_2, fColPT10_2
      ];

      maleView = CollectionViewSource.GetDefaultView(Males); 
      femaleView = CollectionViewSource.GetDefaultView(Females);
      dgMale.MouseDoubleClick += Cell_DoubleClick;
      dgFemale.MouseDoubleClick += Cell_DoubleClick;
    }

    private void Window_KeyDown(object sender, KeyEventArgs e) {
      if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.F) {        
        if (MainTab.SelectedItem == tabStudents) {
          controlsSearchBox.Focus();
        }
      }
    }

    private async void ShowError(string message) {
      if (_isDialogOpen) return; 
      _isDialogOpen = true;
      if (_errorDialog != null && _errorDialog.IsLoaded) return;
      _errorDialog.Content = message;
      await _errorDialog.ShowAsync();
      _isDialogOpen = false;
    }

    private void TextExam1_LostFocus(object sender, RoutedEventArgs e) {
      if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
        if (!int.TryParse(tb.Text, out _)) return;
      if (tb.Text == HighestExam_1) return;
      SetLoading();
      

      var res = _workbookService.SetHighestScore(0, tb.Text, ScoreType.Exam, true);
      if (res.IsFailure) {
        ShowError(res.Error);
        SetLoading(false);
        return;
      }
      HighestExam_1 = tb.Text;
      SetColumnsToHighestScores();
      SetLoading(false);
    }

    private void HighestExam1_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
          if (!int.TryParse(tb.Text, out _)) return;
        if (tb.Text == HighestExam_1) return;
        SetLoading();
        
        var res = _workbookService.SetHighestScore(0, tb.Text, ScoreType.Exam, true); if (res.IsFailure) {
          ShowError(res.Error);
          SetLoading(false);
          e.Handled = true; 
          return;
        }        
        HighestExam_1 = tb.Text;
        SetColumnsToHighestScores();
        SetLoading(false);
        e.Handled = true;
      }
    }
    private void TextExam2_LostFocus(object sender, RoutedEventArgs e) {
      if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
        if (!int.TryParse(tb.Text, out _)) return;
      if (tb.Text == HighestExam_2) return;
      SetLoading();
      

      var res = _workbookService.SetHighestScore(0, tb.Text, ScoreType.Exam, false);
      if (res.IsFailure) {
        ShowError(res.Error);
        SetLoading(false);
        return;
      }
      HighestExam_2 = tb.Text;
      SetColumnsToHighestScores();
      SetLoading(false);
    }

    private void HighestExam2_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
          if (!int.TryParse(tb.Text, out _)) return;
        if (tb.Text == HighestExam_2) return;
        SetLoading();
        
        var res = _workbookService.SetHighestScore(0, tb.Text, ScoreType.Exam, false); if (res.IsFailure) {
          ShowError(res.Error);
          SetLoading(false);
          e.Handled = true; 
          return;
        }        
        HighestExam_2 = tb.Text;
        SetColumnsToHighestScores();
        SetLoading(false);
        e.Handled = true;
      }
    }
    private void TextBoxWW1_LostFocus(object sender, RoutedEventArgs e) {
      if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
        if (!int.TryParse(tb.Text, out _)) return;
      int index = (int)tb.Tag;
      if (tb.Text == HighestWW_1[index]) return;

      SetLoading();
      var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.WrittenWorks, true);
      if (res.IsFailure) {
        ShowError(res.Error);
        SetLoading(false);
        return;
      }
      HighestWW_1[index] = tb.Text;
      SetColumnsToHighestScores();
      SetLoading(false);
    }

    private void HighestScoreWW1_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
          if (!int.TryParse(tb.Text, out _)) return;
        int index = (int)tb.Tag;
        if (tb.Text == HighestWW_1[index]) return;
        SetLoading();

        var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.WrittenWorks, true); if (res.IsFailure) {
          ShowError(res.Error);
          SetLoading(false);
          e.Handled = true; 
          return;
        }        
        HighestWW_1[index] = tb.Text;
        SetColumnsToHighestScores();
        SetLoading(false);
        e.Handled = true;
      }
    }
    private void TextBoxWW2_LostFocus(object sender, RoutedEventArgs e) {
      if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
        if (!int.TryParse(tb.Text, out _)) return;
      int index = (int)tb.Tag;
      if (tb.Text == HighestWW_2[index]) return;

      SetLoading();
      var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.WrittenWorks, false);
      if (res.IsFailure) {
        ShowError(res.Error);
        SetLoading(false);
        return;
      }
      HighestWW_2[index] = tb.Text;
      SetColumnsToHighestScores();
      SetLoading(false);
    }

    private void HighestScoreWW2_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
          if (!int.TryParse(tb.Text, out _)) return;
        int index = (int)tb.Tag;
        if (tb.Text == HighestWW_2[index]) return;
        SetLoading();
        var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.WrittenWorks, false); if (res.IsFailure) {
          ShowError(res.Error);
          SetLoading(false);
          e.Handled = true; 
          return;
        }        
        HighestWW_2[index] = tb.Text;
        SetColumnsToHighestScores();
        SetLoading(false);
        e.Handled = true;
      }
    }
    private void TextBoxPT1_LostFocus(object sender, RoutedEventArgs e) {
      if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
        if (!int.TryParse(tb.Text, out _)) return;
      int index = (int)tb.Tag;
      if (tb.Text == HighestPT_1[index]) return;

      SetLoading();
      var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.PerformanceTasks, Quarter1);
      if (res.IsFailure) {
        ShowError(res.Error);
        SetLoading(false);
        return;
      }
      HighestPT_1[index] = tb.Text;
      SetColumnsToHighestScores();
      SetLoading(false);
    }

    private void HighestScorePT1_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
          if (!int.TryParse(tb.Text, out _)) return;
        int index = (int)tb.Tag;
        if (tb.Text == HighestPT_1[index]) return;
        SetLoading();
        var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.PerformanceTasks, Quarter1); if (res.IsFailure) {
          ShowError(res.Error);
          SetLoading(false);
          e.Handled = true; 
          return;
        }        
        HighestPT_1[index] = tb.Text;
        SetColumnsToHighestScores();
        SetLoading(false);
        e.Handled = true;
      }
    }
    private void TextBoxPT2_LostFocus(object sender, RoutedEventArgs e) {
      if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
        if (!int.TryParse(tb.Text, out _)) return;
      int index = (int)tb.Tag;
      if (tb.Text == HighestPT_2[index]) return;

      SetLoading();
      var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.PerformanceTasks, false);
      if (res.IsFailure) {
        ShowError(res.Error);
        SetLoading(false);
        return;
      }
      HighestPT_2[index] = tb.Text;
      SetColumnsToHighestScores();
      SetLoading(false);
    }

    private void HighestScorePT2_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        if (sender is not TextBox tb) return;
          
      if (!string.IsNullOrEmpty(tb.Text))
          if (!int.TryParse(tb.Text, out _)) return;
        int index = (int)tb.Tag;
        if (tb.Text == HighestPT_2[index]) return;
        SetLoading();
        var res = _workbookService.SetHighestScore(index + 1, tb.Text, ScoreType.PerformanceTasks, false); if (res.IsFailure) {
          ShowError(res.Error);
          SetLoading(false);
          e.Handled = true; 
          return;
        }        
        HighestPT_2[index] = tb.Text;
        SetColumnsToHighestScores();
        SetLoading(false);
        e.Handled = true;
      }
    }

    private void TxtSearchAll_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args) {        
      if (args.Reason != AutoSuggestionBoxTextChangeReason.UserInput)
        return;

      string filterText = sender.Text;

      maleView.Filter = item =>
      {
          if (string.IsNullOrEmpty(filterText)) return true;
          if (item is not StudentWithScores student || string.IsNullOrEmpty(student.Name)) return false;
          return student.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0;
      };
      maleView.Refresh();

      femaleView.Filter = item =>
      {
          if (string.IsNullOrEmpty(filterText)) return true;
          if (item is not StudentWithScores student || string.IsNullOrEmpty(student.Name)) return false;
          return student.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0;
      };
      femaleView.Refresh();
    }


    /* private void TxtSearchAll_TextChanged(object sender, TextChangedEventArgs e) {
      var tb = sender as TextBox; 
      if (tb == null) return;
      string filterText = tb.Text; 
      maleView.Filter = item => {
        if (string.IsNullOrEmpty(filterText)) return true; 
        var student = item as StudentWithScores; 
        if (student == null || string.IsNullOrEmpty(student.Name)) return false; 
        return student.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0;
      }; 
      maleView.Refresh(); 
      femaleView.Filter = item => {
        if (string.IsNullOrEmpty(filterText)) return true; 
        var student = item as StudentWithScores; 
        if (student == null || string.IsNullOrEmpty(student.Name)) return false; 
        return student.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0;
      }; 
      femaleView.Refresh(); 
    } */

    private void NumberOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e) => 
      e.Handled = !int.TryParse(e.Text, out _);    

    private void GradeSwitch_Toggled(object sender, RoutedEventArgs e) { 
      var toggle = sender as ToggleSwitch; 
      if (toggle == null) return; 

      
      ShowGrades = toggle.IsOn;
      
      SafeVisibilityToggle(mColGrade_1, true, true);
      SafeVisibilityToggle(fColGrade_1, true, true);
      SafeVisibilityToggle(mColGrade_2, false, true);
      SafeVisibilityToggle(fColGrade_2, false, true);
    }

    private void QuarterSwitch_Toggled(object sender, RoutedEventArgs e) { 
      var toggle = sender as ToggleSwitch; 
      if (toggle == null) return; 

      
      Quarter1 = toggle.IsOn;
      SetGridColumnsVisibility();          
    }

    private void Cell_DoubleClick(object sender, MouseButtonEventArgs e) {
      if (e.OriginalSource is FrameworkElement fe && fe.Parent is DataGridCell cell) {        
        
        var column = cell?.Column;
        var rowData = cell?.DataContext;

        if (column == null || rowData == null) return;

        if (column.Header.ToString() == "Male" || column.Header.ToString() == "Female") {
          ShowError($"Double-clicked Age cell: {((StudentWithScores)rowData).Name}");
        }
      }      
    } 


    private void MyDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e) {
      if (e.EditingElement is not  TextBox tb) return;
      bool success = int.TryParse(tb.Text, out int newValue);

      if (!success) return;

      if (newValue < 0) return;


      var rowData = e.Row.Item;
      var binding = (e.Column as DataGridBoundColumn)?.Binding as Binding; 

      SetLoading(true);
      if (binding != null) { 
        var propertyName = binding.Path.Path; 
        var propInfo = rowData.GetType().GetProperty(propertyName); 
        var oldValue = propInfo?.GetValue(rowData); 

        if (newValue.ToString() == oldValue?.ToString()) {
          SetLoading(false);
          return;
        }

        bool successHeader = int.TryParse(e.Column.Header?.ToString() ?? "", out int header);         

        if (newValue > header || !successHeader) { 
          tb.Text = oldValue?.ToString(); 
          ShowError("Value must be less than the Highest Value"); 
          SetLoading(false);
          return; 
        }

        string scoreType = propertyName.Substring(0, 2);

        int underscoreIndex = propertyName.IndexOf('_'); 
        string quarter = propertyName.Substring(underscoreIndex + 1);
        bool numberSucces = int.TryParse(propertyName.Substring(underscoreIndex - 1, 1), out int number);
        
        if (scoreType == "WW" && numberSucces) {
          var res = _workbookService.SetStudentScore(((StudentWithScores)rowData).Row, number, tb.Text, ScoreType.WrittenWorks, quarter == "1");
          if (res.IsFailure) {
            tb.Text = oldValue?.ToString(); 
            ShowError(res.Error);
          }
        } else if (scoreType == "PT" && numberSucces) {
          var res = _workbookService.SetStudentScore(((StudentWithScores)rowData).Row, number, tb.Text, ScoreType.PerformanceTasks, quarter == "1");
          if (res.IsFailure) {
            tb.Text = oldValue?.ToString(); 
            ShowError(res.Error);
          }
        } else if (scoreType == "EX" && !numberSucces) {
          var res = _workbookService.SetStudentScore(((StudentWithScores)rowData).Row, 0, tb.Text, ScoreType.Exam, quarter == "1");
          if (res.IsFailure) {
            tb.Text = oldValue?.ToString(); 
            ShowError(res.Error);
          }
        }
        rowData.GetType().GetProperty(propertyName)?.SetValue(rowData, tb.Text);
        
        ComputeTransmutedScores(rowData as StudentWithScores);
        // ShowError($"Old value: {oldValue}, New value: {newValue}, Property: {propertyName}, Property Type: {propInfo?.PropertyType}"); 
      }

      e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
      SetLoading(false);
    }


    private async void File_Click(object sender, RoutedEventArgs e) {

      Males.Clear();
      Females.Clear();
      var dialog = new OpenFileDialog {
        Title = "Select an Excel file",
        Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
      };

      SetLoading(true);
      if (dialog.ShowDialog() == true) {
          
        Result resultFile = await Task.Run(() => {

          var isECR = _workbookService.IsFileECR(dialog.FileName);
          if (isECR.IsFailure) {
            return Result.Failure(isECR.Error);
          }

          var res = _workbookService.GetBounds();
          if (res.IsFailure) {
            return Result.Failure(res.Error);
          }

          return Result.Success();
        });


        if (resultFile.IsFailure) {
          ShowError(resultFile.Error);
          SetLoading(false);
          return;
        }


        HighestWW_1.Clear();
        HighestPT_1.Clear();
        HighestWW_2.Clear();
        HighestPT_2.Clear();

        await foreach (var result in _workbookService.ReadScoresAsync()) {

          if (result.IsFailure) {
            ShowError(result.Error);
            break;
          }

          var (
            highestWW_1, 
            highestPT_1,
            exam_1, 
            highestWW_2, 
            highestPT_2, 
            exam_2, 
            maleWithScores, 
            femaleWithScores
            ) = result.Value;

          await Application.Current.Dispatcher.BeginInvoke(() => {
            
            HighestExam_1 = exam_1;
            HighestExam_2 = exam_2;
            
            foreach (var hs1 in highestWW_1)
              HighestWW_1.Add(hs1);

            foreach (var hs2 in highestWW_2)
              HighestWW_2.Add(hs2);

            foreach (var hs1 in highestPT_1)
              HighestPT_1.Add(hs1);
            

            foreach (var hs2 in highestPT_2) 
              HighestPT_2.Add(hs2);
            

            foreach (var m in maleWithScores) 
              Males.Add(m);

            foreach (var f in femaleWithScores) 
              Females.Add(f);
          });

        }

        foreach (var s in Males) 
          ComputeTransmutedScores(s);
        
        foreach (var s in Females) 
          ComputeTransmutedScores(s);

        tsHSQuarter.OnContent = _workbookService.FirstSem ? "1st" : "3rd";
        tsSQuarter.OnContent = _workbookService.FirstSem ? "1st" : "3rd";
        tsHSQuarter.OffContent = _workbookService.FirstSem ? "2nd" : "4th";
        tsSQuarter.OffContent = _workbookService.FirstSem ? "2nd" : "4th";

        txtFileName.Text = System.IO.Path.GetFileNameWithoutExtension(_workbookService.FilePath);

        tabHighestScores.Visibility = Visibility.Visible;
        tabStudents.Visibility = Visibility.Visible;
        tabHighestScores.IsEnabled = true;
        tabStudents.IsEnabled = true;

        SetColumnsToHighestScores();
      }
      SetLoading(false);
    }

    private void SetGridColumnsVisibility() {
      
      for (int i = 0; i < mWWSet1Columns.Length; i++) {
        mWWSet1Columns[i].Visibility = Quarter1 && !string.IsNullOrWhiteSpace(HighestWW_1[i])
            ? Visibility.Visible
            : Visibility.Collapsed;
        fWWSet1Columns[i].Visibility = Quarter1 && !string.IsNullOrWhiteSpace(HighestWW_1[i])
            ? Visibility.Visible
            : Visibility.Collapsed;          
      }      

      for (int i = 0; i < mPTSet1Columns.Length; i++) {
        mPTSet1Columns[i].Visibility = Quarter1 && !string.IsNullOrWhiteSpace(HighestPT_1[i])
            ? Visibility.Visible
            : Visibility.Collapsed;
        fPTSet1Columns[i].Visibility = Quarter1 && !string.IsNullOrWhiteSpace(HighestPT_1[i])
            ? Visibility.Visible
            : Visibility.Collapsed;
      }


      for (int i = 0; i < mWWSet2Columns.Length; i++) {
        mWWSet2Columns[i].Visibility = !Quarter1 && !string.IsNullOrWhiteSpace(HighestWW_2[i])
            ? Visibility.Visible
            : Visibility.Collapsed;
        fWWSet2Columns[i].Visibility = !Quarter1 && !string.IsNullOrWhiteSpace(HighestWW_2[i])
            ? Visibility.Visible
            : Visibility.Collapsed;
      }


      for (int i = 0; i < mPTSet2Columns.Length; i++) {
        mPTSet2Columns[i].Visibility = !Quarter1 && !string.IsNullOrWhiteSpace(HighestPT_2[i])
            ? Visibility.Visible
            : Visibility.Collapsed;
        fPTSet2Columns[i].Visibility = !Quarter1 && !string.IsNullOrWhiteSpace(HighestPT_2[i])
            ? Visibility.Visible
            : Visibility.Collapsed;
      }

      if (!string.IsNullOrWhiteSpace(HighestExam_1)) SafeVisibilityToggle(mColExam_1);
      SafeVisibilityToggle(mE1);
      if (!string.IsNullOrWhiteSpace(HighestExam_2)) SafeVisibilityToggle(mColExam_2, false);
      SafeVisibilityToggle(mE2, false);
      if (!string.IsNullOrWhiteSpace(HighestExam_1)) SafeVisibilityToggle(fColExam_1);
      SafeVisibilityToggle(fE1);
      if (!string.IsNullOrWhiteSpace(HighestExam_2)) SafeVisibilityToggle(fColExam_2, false);
      SafeVisibilityToggle(fE2, false);

      SafeVisibilityToggle(mColGrade_1, true, true);
      SafeVisibilityToggle(mColGrade_2, false, true);
      SafeVisibilityToggle(fColGrade_1, true, true);
      SafeVisibilityToggle(fColGrade_2, false, true);

      SafeVisibilityToggle(mWW1);
      SafeVisibilityToggle(mPT1);
      SafeVisibilityToggle(mWW2, false);
      SafeVisibilityToggle(mPT2, false);
      SafeVisibilityToggle(fWW1);
      SafeVisibilityToggle(fPT1);
      SafeVisibilityToggle(fWW2, false);
      SafeVisibilityToggle(fPT2, false);

      SafeVisibilityToggle(bdHighestWW_1);
      SafeVisibilityToggle(bdHighestPT_1);
      SafeVisibilityToggle(bdHighestWW_2, false);
      SafeVisibilityToggle(bdHighestPT_2, false);
      
    }

    private void ComputeTransmutedScores(StudentWithScores s) {
      s.Transmuted_1 = _workbookService.GetComputedGrade(
        HighestWW_1.ToList(), 
        HighestPT_1.ToList(),
        HighestExam_1,
        [s.WW1_1, s.WW2_1, s.WW3_1, s.WW4_1, s.WW5_1, s.WW6_1, s.WW7_1, s.WW8_1, s.WW9_1, s.WW10_1],
        [s.PT1_1, s.PT2_1, s.PT3_1, s.PT4_1, s.PT5_1, s.PT6_1, s.PT7_1, s.PT8_1, s.PT9_1, s.PT10_1],
        s.EX_1
      ).ToString();
      s.Transmuted_2 = _workbookService.GetComputedGrade(
        HighestWW_2.ToList(), 
        HighestPT_2.ToList(),
        HighestExam_2,
        [s.WW1_2, s.WW2_2, s.WW3_2, s.WW4_2, s.WW5_2, s.WW6_2, s.WW7_2, s.WW8_2, s.WW9_2, s.WW10_2],
        [s.PT1_2, s.PT2_2, s.PT3_2, s.PT4_2, s.PT5_2, s.PT6_2, s.PT7_2, s.PT8_2, s.PT9_2, s.PT10_2],
        s.EX_2
      ).ToString();
    }

    private void SetColumnsToHighestScores() {
      for (int i = 0; i < mWWSet1Columns.Length; i++) {
        string headerValue = HighestWW_1[i];
        mWWSet1Columns[i].Header = headerValue;
        fWWSet1Columns[i].Header = headerValue;
      
        ApplyThresholdStyle(mWWSet1Columns[i], headerValue, "WW" + (i+1) + "_1");
        ApplyThresholdStyle(fWWSet1Columns[i], headerValue, "WW" + (i+1) + "_1");
      }

      for (int i = 0; i < mPTSet1Columns.Length; i++) {
        string headerValue = HighestPT_1[i]; 
        mPTSet1Columns[i].Header = headerValue; 
        fPTSet1Columns[i].Header = headerValue; 
        
        ApplyThresholdStyle(mPTSet1Columns[i], headerValue, "PT" + (i+1) + "_1"); 
        ApplyThresholdStyle(fPTSet1Columns[i], headerValue, "PT" + (i+1) + "_1");
      }

      for (int i = 0; i < mWWSet2Columns.Length; i++) {
        string headerValue = HighestWW_2[i];
        mWWSet2Columns[i].Header = headerValue;
        fWWSet2Columns[i].Header = headerValue;

        ApplyThresholdStyle(mWWSet2Columns[i], headerValue, "WW" + (i+1) + "_2");
        ApplyThresholdStyle(fWWSet2Columns[i], headerValue, "WW" + (i+1) + "_2");
      }

      for (int i = 0; i < mPTSet2Columns.Length; i++) {
        string headerValue = HighestPT_2[i];
        mPTSet2Columns[i].Header = headerValue;
        fPTSet2Columns[i].Header = headerValue;

        ApplyThresholdStyle(mPTSet2Columns[i], headerValue, "PT" + (i+1) + "_2");
        ApplyThresholdStyle(fPTSet2Columns[i], headerValue, "PT" + (i+1) + "_2");
      }
      ApplyThresholdStyle(mColExam_1, HighestExam_1, "EX_1");
      ApplyThresholdStyle(fColExam_1, HighestExam_1, "EX_1");
      ApplyThresholdStyle(mColExam_2, HighestExam_2, "EX_2");
      ApplyThresholdStyle(fColExam_2, HighestExam_2, "EX_2");       
      SetGridColumnsVisibility();
    }

    private static void ApplyThresholdStyle(DataGridTextColumn column, string headerValue, string bindingPath) {
      var baseStyle = (Style)Application.Current.Resources["ThresholdCellStyle"];
      var style = new Style(typeof(DataGridCell), baseStyle);
      style.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Center));
      style.Setters.Add(new Setter(DataGridCell.BackgroundProperty,
          new Binding(bindingPath)
          {
              Converter = new LessThanPercentageOfHeaderConverter(),
              ConverterParameter = headerValue
          }));

      // style.EventSetters.Add(new EventSetter(DataGridCell.MouseDoubleClickEvent, new MouseButtonEventHandler(Cell_DoubleClick)));    

      column.CellStyle = style;
    }

    private void SafeVisibilityToggle(DataGridColumn column, bool quarter1 = true, bool showGrades = false) {
      if (column == null) return;
      if (quarter1) {
        if (showGrades)
          column.Visibility = ShowGrades && Quarter1 ? Visibility.Visible : Visibility.Collapsed;
        else
          column.Visibility = Quarter1 ? Visibility.Visible : Visibility.Collapsed;
      } else {
        if (showGrades)
          column.Visibility = ShowGrades && !Quarter1 ? Visibility.Visible : Visibility.Collapsed;
        else
          column.Visibility = !Quarter1 ? Visibility.Visible : Visibility.Collapsed;
      }
    } 
    private void SafeVisibilityToggle(FrameworkElement el, bool quarter = true) {
      if (el == null) return;
      if (quarter)
        el.Visibility = Quarter1 ? Visibility.Visible : Visibility.Collapsed;
      else
        el.Visibility = !Quarter1 ? Visibility.Visible : Visibility.Collapsed;
    } 

    private void SetLoading(bool isLoading = true) {
      this.IsEnabled = !isLoading;
      progressBarMain.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
    }
  }
}