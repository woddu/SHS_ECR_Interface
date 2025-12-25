using CSharpFunctionalExtensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;

namespace WpfSample;

public class WorkbookService {
  private record GradeRange(
    double Min,
    double Max,
    int Transmuted
  );

  private readonly static List<GradeRange> gradeTable = [
    new GradeRange (100.00, 100.00, 100 ),
    new GradeRange (98.40, 99.99, 99 ),
    new GradeRange (96.80, 98.39, 98 ),
    new GradeRange (95.20, 96.79, 97 ),
    new GradeRange (93.60, 95.19, 96 ),
    new GradeRange (92.00, 93.59, 95 ),
    new GradeRange (90.40, 91.99, 94 ),
    new GradeRange (88.80, 90.39, 93 ),
    new GradeRange (87.20, 88.79, 92 ),
    new GradeRange (85.60, 87.19, 91 ),
    new GradeRange (84.00, 85.59, 90 ),
    new GradeRange (82.40, 83.99, 89 ),
    new GradeRange (80.80, 82.39, 88 ),
    new GradeRange (79.20, 80.79, 87 ),
    new GradeRange (77.60, 79.19, 86 ),
    new GradeRange (76.00, 77.59, 85 ),
    new GradeRange (74.40, 75.99, 84 ),
    new GradeRange (72.80, 74.39, 83 ),
    new GradeRange (71.20, 72.79, 82 ),
    new GradeRange (69.60, 71.19, 81 ),
    new GradeRange (68.00, 69.59, 80 ),
    new GradeRange (66.40, 67.99, 79 ),
    new GradeRange (64.80, 66.39, 78 ),
    new GradeRange (63.20, 64.79, 77 ),
    new GradeRange (61.60, 63.19, 76 ),
    new GradeRange (60.00, 61.59, 75 ),
    new GradeRange (56.00, 59.99, 74 ),
    new GradeRange (52.00, 55.99, 73 ),
    new GradeRange (48.00, 51.99, 72 ),
    new GradeRange (44.00, 47.99, 71 ),
    new GradeRange (40.00, 43.99, 70 ),
    new GradeRange (36.00, 39.99, 69 ),
    new GradeRange (32.00, 35.99, 68 ),
    new GradeRange (28.00, 31.99, 67 ),
    new GradeRange (24.00, 27.99, 66 ),
    new GradeRange (20.00, 23.99, 65 ),
    new GradeRange (16.00, 19.99, 64 ),
    new GradeRange (12.00, 15.99, 63 ),
    new GradeRange (8.00,  11.99, 62 ),
    new GradeRange (4.00,  7.99,  61 ),
    new GradeRange (0.00,  3.99,  60 )
  ];

  private string[] _sharedStrings;

  private static uint _maleScoresStartRow_1 = 0u;
  private static uint _femaleScoresStartRow_1 = 0u;
  private static uint _maleScoresStartRow_2 = 0u;
  private static uint _femaleScoresStartRow_2 = 0u;

  public readonly static string[] tracks = [
    "Core Subject (All Tracks)",
    "Academic Track (except Immersion)",
    "Work Immersion/ Culminating Activity (for Academic Track)",
    "TVL/ Sports/ Arts and Design Track"
  ];

  public readonly static double[,] weightedScores = {
    { 0.25, 0.50, 0.25 }, // Core Subject (All Tracks)
    { 0.25, 0.45, 0.30 }, // Academic Track (except Immersion)
    { 0.35, 0.40, 0.25 }, // Work Immersion/ Culminating Activity (for Academic Track)
    { 0.20, 0.60, 0.20 }  // TVL/ Sports/ Arts and Design Track
  };

  private static readonly List<string> requiredSheetNames1 = [
    "INPUT DATA",
    "1ST",
    "2ND",
    "Final Semestral Grade"
  ];

  private static readonly List<string> requiredSheetNames2 = [
    "INPUT DATA",
    "3RD",
    "4TH",
    "Final Semestral Grade"
  ];

  public bool FirstSem { get; private set; } = true;

  public string FilePath { get; private set; } = "";

  public string Track { get; private set; } = "";

  public double[] WeightedScores { get; private set; } = { weightedScores[0, 0], weightedScores[0, 1], weightedScores[0, 2] };  
  
  public Result IsFileECR(string path) {

    FilePath = path;


    using var doc = new ECRSpreadSheetDocument();
    var openResult = doc.Open(FilePath);
    if (openResult.IsFailure)
      return Result.Failure(openResult.Error);

    if (doc.Sheets.IsFailure)
      return Result.Failure(doc.Sheets.Error);

    var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (Sheet s in doc.Sheets.Value)
      sheetNames.Add(s.Name);

    if (doc.WorkbookPart.IsFailure)
      return Result.Failure(doc.WorkbookPart.Error);
    
    var sst = doc.WorkbookPart.Value.SharedStringTablePart?.SharedStringTable;
    _sharedStrings = sst != null
        ? sst.Elements<SharedStringItem>().Select(i => i.InnerText).ToArray()
        : Array.Empty<string>();


    var missingFromSet1 = requiredSheetNames1.Where(req => !sheetNames.Contains(req));
    var missingFromSet2 = requiredSheetNames2.Where(req => !sheetNames.Contains(req));
    var missingSheets1 = missingFromSet1.Union(missingFromSet2);
      
    bool matchesSetA = requiredSheetNames1.All(sheetNames.Contains);

    bool matchesSetB = requiredSheetNames2.All(sheetNames.Contains);

    if (!matchesSetA && matchesSetB) {
      FirstSem = false;
    } else if (matchesSetA && !matchesSetB) {
      FirstSem = true;
    } else if (!matchesSetA || !matchesSetB) {
      return Result.Failure("Required sheets not found: " + string.Join(", ", missingSheets1));
    }

    return Result.Success();
  }
  
  public Result GetBounds() {
  
    using var doc = new ECRSpreadSheetDocument();
    var openResult = doc.Open(FilePath);
    if (openResult.IsFailure)
      return Result.Failure(openResult.Error);

    if (doc.Sheets.IsFailure)
      return Result.Failure(doc.Sheets.Error);

    var sheet = doc.SheetByName(FirstSem ? requiredSheetNames1[1] : requiredSheetNames2[1]);
    if (sheet.IsFailure)
      return Result.Failure(sheet.Error);

    var wsPartResult1 = doc.WsPartById(sheet.Value.Id!);
    if (wsPartResult1.IsFailure)
      return Result.Failure(wsPartResult1.Error);

    if (wsPartResult1.Value is not WorksheetPart wsPart1)
      return Result.Failure("WorksheetPart failure");

    if (ReadCellValue(wsPart1, "A12")!.Equals("male", StringComparison.OrdinalIgnoreCase) &&
      ReadCellValue(wsPart1, "A38")!.Equals("female", StringComparison.OrdinalIgnoreCase)) {
      _maleScoresStartRow_1 = 13u;
      _femaleScoresStartRow_1 = 39u;
    } else if (ReadCellValue(wsPart1, "A12")!.Equals("male", StringComparison.OrdinalIgnoreCase) &&
      ReadCellValue(wsPart1, "A63")!.Equals("female", StringComparison.OrdinalIgnoreCase)) {
      _maleScoresStartRow_1 = 13u;
      _femaleScoresStartRow_1 = 64u;
    } else if (ReadCellValue(wsPart1, "A12")!.Equals("male", StringComparison.OrdinalIgnoreCase) &&
      ReadCellValue(wsPart1, "A68")!.Equals("female", StringComparison.OrdinalIgnoreCase)) {
      _maleScoresStartRow_1 = 13u;
      _femaleScoresStartRow_1 = 69u;
    } else {
      var result1 = ScanSheetForBounds(wsPart1, out _maleScoresStartRow_1, out _femaleScoresStartRow_1);
      if (result1.IsFailure)
        return result1;
    }

    // ---------------- SECOND SHEET (names) ----------------
    sheet = doc.SheetByName(FirstSem ? requiredSheetNames1[2] : requiredSheetNames2[2]);
    if (sheet.IsFailure)
      return Result.Failure(sheet.Error);

    var wsPartResult2 = doc.WsPartById(sheet.Value.Id!);
    if (wsPartResult2.IsFailure)
      return Result.Failure(wsPartResult2.Error);

    if (wsPartResult2.Value is not WorksheetPart wsPart2)
      return Result.Failure("WorksheetPart failure");

    if (ReadCellValue(wsPart2, "A12")!.Equals("male", StringComparison.OrdinalIgnoreCase) &&
      ReadCellValue(wsPart2, "A38")!.Equals("female", StringComparison.OrdinalIgnoreCase)) {
      _maleScoresStartRow_2 = 13u;
      _femaleScoresStartRow_2 = 39u;
    } else if (ReadCellValue(wsPart2, "A12")!.Equals("male", StringComparison.OrdinalIgnoreCase) &&
      ReadCellValue(wsPart2, "A63")!.Equals("female", StringComparison.OrdinalIgnoreCase)) {
      _maleScoresStartRow_2 = 13u;
      _femaleScoresStartRow_2 = 64u;
    } else if (ReadCellValue(wsPart2, "A12")!.Equals("male", StringComparison.OrdinalIgnoreCase) &&
      ReadCellValue(wsPart2, "A68")!.Equals("female", StringComparison.OrdinalIgnoreCase)) {
      _maleScoresStartRow_2 = 13u;
      _femaleScoresStartRow_2 = 69u;
    } else {
      var result2 = ScanSheetForBounds(wsPart2, out _maleScoresStartRow_2, out _femaleScoresStartRow_2);
      if (result2.IsFailure)
        return result2;
    }
    return Result.Success();
  
  }
  
  private Result ScanSheetForBounds(
    WorksheetPart wsPart,
    out uint maleStart,
    out uint femaleStart
  ) {
    maleStart = 0;
    femaleStart = 0;

    using var reader = OpenXmlReader.Create(wsPart);

    while (reader.Read()) {
      if (reader.ElementType == typeof(Row)) {
        var row = (Row)reader.LoadCurrentElement();
        uint rowIndex = row.RowIndex;

        // Only look at column A
        var cell = row.Elements<Cell>()
                      .FirstOrDefault(c => c.CellReference?.Value.StartsWith("A") == true);

        if (rowIndex == 8u) {
          var track = row.Elements<Cell>()
                      .FirstOrDefault(c => c.CellReference?.Value.StartsWith("AE") == true);
          Track = GetCellValue(track).Trim();
        }

        if (cell == null)
          continue;

        string val = GetCellValue(cell).Trim();

        if (val.Equals("male", StringComparison.OrdinalIgnoreCase)) {
          maleStart = rowIndex + 1;
        } else if (val.Equals("female", StringComparison.OrdinalIgnoreCase)) {
          femaleStart = rowIndex + 1;
          return Result.Success();
        }
      }
    }

    return Result.Failure("Could not find male/female markers");
  }

  public async IAsyncEnumerable<
    Result<(
      List<string> highestWW_1, 
      List<string> highestPT_1,
      string exam_1, 
      List<string> highestWW_2, 
      List<string> highestPT_2,
      string exam_2, 
      List<StudentWithScores> maleStudents, 
      List<StudentWithScores> femaleStudents
    )>
  > ReadScoresAsync(int chunkSize = 3) {
    string nameColumn = "B";
        
    using var doc = new ECRSpreadSheetDocument();
    var openResult = doc.Open(FilePath);
    if (openResult.IsFailure) {
      yield return Result.Failure<(List<string>, List<string>, string, List<string>, List<string>, string, List<StudentWithScores>, List<StudentWithScores>)>(openResult.Error);
      yield break;
    }

    var sheet1 = doc.SheetByName(FirstSem ? requiredSheetNames1[1] : requiredSheetNames2[1]);
    var sheet2 = doc.SheetByName(FirstSem ? requiredSheetNames1[2] : requiredSheetNames2[2]);

    if (sheet1.IsFailure && sheet2.IsFailure) {
      yield return Result.Failure<(List<string>, List<string>, string, List<string>, List<string>, string, List<StudentWithScores>, List<StudentWithScores>)>("Required sheet not found");
      yield break;
    }

    var wsPartResult1 = doc.WsPartById(sheet1.Value.Id!);
    if (wsPartResult1.IsFailure) {
      yield return Result.Failure<(List<string>, List<string>, string, List<string>, List<string>, string, List<StudentWithScores>, List<StudentWithScores>)>(wsPartResult1.Error);
      yield break;
    }

    var wsPartResult2 = doc.WsPartById(sheet2.Value.Id!);
    if (wsPartResult2.IsFailure) {
      yield return Result.Failure<(List<string>, List<string>, string, List<string>, List<string>, string, List<StudentWithScores>, List<StudentWithScores>)>(wsPartResult2.Error);
      yield break;
    }

    if (wsPartResult1.Value is not WorksheetPart wsPart1) {
      yield return Result.Failure<(List<string>, List<string>, string, List<string>, List<string>, string, List<StudentWithScores>, List<StudentWithScores>)>("WorksheetPart not found");
      yield break;
    }

    if (wsPartResult2.Value is not WorksheetPart wsPart2) {
      yield return Result.Failure<(List<string>, List<string>, string, List<string>, List<string>, string, List<StudentWithScores>, List<StudentWithScores>)>("WorksheetPart not found");
      yield break;
    }

    SheetData sheetData1 = doc.SheetData(wsPart1);
    SheetData sheetData2 = doc.SheetData(wsPart2);

    if (sheetData1 == null && sheetData2 == null) {
      yield return Result.Failure<(List<string>, List<string>, string, List<string>, List<string>, string, List<StudentWithScores>, List<StudentWithScores>)>("SheetData is null");
      yield break;
    }
    
    string[,] Cols = {
      { "F", "O" },   // WrittenWorks_1
      { "S", "AB" }   // PerformanceTasks_1
    };
    string examCol = "AF";

    uint highestScoresRow = 11;

    var highestWWChunk_1 = new List<string>(chunkSize);
    var highestPTChunk_1 = new List<string>(chunkSize);
    string exam_1 = "";
    for (int i = 0; i < Cols.GetLength(0); i++) {
        await Task.Yield();
        int startColIdx = ColNameToNumber(Cols[i, 0]);
        int endColIdx   = ColNameToNumber(Cols[i, 1]);

        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
            string colName = ColNumberToName(colIdx);
            string cellRef = $"{colName}{highestScoresRow}";
            Cell cell = sheetData1.Descendants<Cell>()
                                  .FirstOrDefault(c => c.CellReference == cellRef)!;
            string val = GetCellValue(cell);
          
            if (i == 0) 
              highestWWChunk_1.Add(val); 
            else {
              highestPTChunk_1.Add(val);
            }
        }
        if (highestWWChunk_1.Count >= chunkSize || highestPTChunk_1.Count >= chunkSize) {
            yield return Result.Success((highestWWChunk_1, highestPTChunk_1, exam_1, new List<string>(), new List<string>(), "", new List<StudentWithScores>(), new List<StudentWithScores>()));
            highestWWChunk_1 = new List<string>(chunkSize);
            highestPTChunk_1 = new List<string>(chunkSize);
        }
    }
    Cell highestExamCell = sheetData1.Descendants<Cell>()
                                  .FirstOrDefault(c => c.CellReference == $"{examCol}11")!;
    exam_1 = GetCellValue(highestExamCell);
    if (highestWWChunk_1.Count > 0 || highestPTChunk_1.Count > 0) {
        yield return Result.Success((highestWWChunk_1, highestPTChunk_1, exam_1, new List<string>(), new List<string>(),  "", new List<StudentWithScores>(), new List<StudentWithScores>()));
    }

    var highestWWChunk_2 = new List<string>(chunkSize);
    var highestPTChunk_2 = new List<string>(chunkSize);
    string exam_2 = "";
    for (int i = 0; i < Cols.GetLength(0); i++) {
        await Task.Yield();
        int startColIdx = ColNameToNumber(Cols[i, 0]);
        int endColIdx   = ColNameToNumber(Cols[i, 1]);

        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
            string colName = ColNumberToName(colIdx);
            string cellRef = $"{colName}{highestScoresRow}";
            Cell cell = sheetData2.Descendants<Cell>()
                                  .FirstOrDefault(c => c.CellReference == cellRef)!;
            string val = GetCellValue(cell);
          
            if (i == 0) 
              highestWWChunk_2.Add(val); 
            else
              highestPTChunk_2.Add(val); 
        }
        if (highestWWChunk_2.Count >= chunkSize || highestPTChunk_2.Count >= chunkSize) {            
            yield return Result.Success((new List<string>(), new List<string>(), exam_1, highestWWChunk_2, highestPTChunk_2, exam_2, new List<StudentWithScores>(), new List<StudentWithScores>()));
            highestWWChunk_2 = new List<string>(chunkSize);
            highestPTChunk_2 = new List<string>(chunkSize);
        }
    }
    highestExamCell = sheetData2.Descendants<Cell>()
                                  .FirstOrDefault(c => c.CellReference == $"{examCol}11")!;
    exam_2 = GetCellValue(highestExamCell);
    if (highestWWChunk_2.Count > 0 || highestPTChunk_2.Count > 0) {
        yield return Result.Success((new List<string>(), new List<string>(), exam_1, highestWWChunk_2, highestPTChunk_2, exam_2, new List<StudentWithScores>(), new List<StudentWithScores>()));
    }


    var maleEndRow = _femaleScoresStartRow_1 - 2;
    var femaleEndRow = _femaleScoresStartRow_1 + 70;
    // --- Male students ---
    var maleChunk = new List<StudentWithScores>(chunkSize);    

    for (uint rowIndex = _maleScoresStartRow_1; rowIndex <= maleEndRow; rowIndex++) {
      await Task.Yield();

      string nameCellRef = $"{nameColumn}{rowIndex}";
      Cell nameCell = sheetData1.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == nameCellRef)!;
      string nameVal = GetCellValue(nameCell);

      if (!string.IsNullOrWhiteSpace(nameVal) && nameVal.Trim() != "0") {
        var student = new StudentWithScores(nameVal);
        student.Row = rowIndex;
        // WrittenWorks_1
        int startColIdx = ColNameToNumber(Cols[0, 0]);
        int endColIdx = ColNameToNumber(Cols[0, 1]);
        int wwCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, wwCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex}";
          Cell cell = sheetData1.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);
          
          student.GetType().GetProperty($"WW{wwCounter}_1")?.SetValue(student, val);
        }

        // PerformanceTasks_1
        startColIdx = ColNameToNumber(Cols[1, 0]);
        endColIdx = ColNameToNumber(Cols[1, 1]);
        int ptCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, ptCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex}";
          Cell cell = sheetData1.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);

          student.GetType().GetProperty($"PT{ptCounter}_1")?.SetValue(student, val);
        }
        
        // Exam_1
        string examCellRef = $"{examCol}{rowIndex}";
        examCellRef = $"{examCol}{rowIndex - _maleScoresStartRow_1 + _maleScoresStartRow_2}";
        Cell examCell = sheetData1.Descendants<Cell>()
                                  .FirstOrDefault(c => c.CellReference == examCellRef)!;
        student.EX_1 = GetCellValue(examCell);

        // WrittenWorks_2
        startColIdx = ColNameToNumber(Cols[0, 0]);
        endColIdx = ColNameToNumber(Cols[0, 1]);
        wwCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, wwCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex - _maleScoresStartRow_1 + _maleScoresStartRow_2}";
          Cell cell = sheetData2.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);

          student.GetType().GetProperty($"WW{wwCounter}_2")?.SetValue(student, val);
        }

        // PerformanceTasks_2
        startColIdx = ColNameToNumber(Cols[1, 0]);
        endColIdx = ColNameToNumber(Cols[1, 1]);
        ptCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, ptCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex - _maleScoresStartRow_1 + _maleScoresStartRow_2}";
          Cell cell = sheetData2.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);

          student.GetType().GetProperty($"PT{ptCounter}_2")?.SetValue(student, val);
        }

        // Exam_2
        examCellRef = $"{examCol}{rowIndex - _maleScoresStartRow_1 + _maleScoresStartRow_2}";
        examCell = sheetData2.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == examCellRef)!;
        student.EX_2 = GetCellValue(examCell);

        maleChunk.Add(student);        
      }


      if (maleChunk.Count >= chunkSize) {
        yield return Result.Success((new List<string>(), new List<string>(), exam_1, new List<string>(), new List<string>(), exam_2, maleChunk, new List<StudentWithScores>()));
        maleChunk = new List<StudentWithScores>(chunkSize);
      }
    }

    if (maleChunk.Count > 0)
      yield return Result.Success((new List<string>(), new List<string>(), exam_1, new List<string>(), new List<string>(), exam_2, maleChunk, new List<StudentWithScores>()));

    // --- Female students ---
    var femaleChunk = new List<StudentWithScores>(chunkSize);    

    for (uint rowIndex = _femaleScoresStartRow_1; rowIndex <= femaleEndRow; rowIndex++) {
      await Task.Yield();

      string nameCellRef = $"{nameColumn}{rowIndex}";
      Cell nameCell = sheetData1.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == nameCellRef)!;
      string nameVal = GetCellValue(nameCell);

      if (!string.IsNullOrWhiteSpace(nameVal) && nameVal.Trim() != "0") {
        var student = new StudentWithScores(nameVal);
        student.Row = rowIndex;
        // WrittenWorks_1
        int startColIdx = ColNameToNumber(Cols[0, 0]);
        int endColIdx = ColNameToNumber(Cols[0, 1]);
        int wwCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, wwCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex}";
          Cell cell = sheetData1.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);

          student.GetType().GetProperty($"WW{wwCounter}_1")?.SetValue(student, val);
        }

        // PerformanceTasks_1
        startColIdx = ColNameToNumber(Cols[1, 0]);
        endColIdx = ColNameToNumber(Cols[1, 1]);
        int ptCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, ptCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex}";
          Cell cell = sheetData1.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);

          student.GetType().GetProperty($"PT{ptCounter}_1")?.SetValue(student, val);
        }

        // Exam_1
        string examCellRef = $"{examCol}{rowIndex}";
        Cell examCell = sheetData1.Descendants<Cell>()
                                  .FirstOrDefault(c => c.CellReference == examCellRef)!;
        student.EX_1 = GetCellValue(examCell);

        // WrittenWorks_2
        startColIdx = ColNameToNumber(Cols[0, 0]);
        endColIdx = ColNameToNumber(Cols[0, 1]);
        wwCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, wwCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex - _femaleScoresStartRow_1 + _femaleScoresStartRow_2}";
          Cell cell = sheetData2.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);

          student.GetType().GetProperty($"WW{wwCounter}_2")?.SetValue(student, val);
        }

        // PerformanceTasks_2
        startColIdx = ColNameToNumber(Cols[1, 0]);
        endColIdx = ColNameToNumber(Cols[1, 1]);
        ptCounter = 1;
        for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++, ptCounter++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{rowIndex - _femaleScoresStartRow_1 + _femaleScoresStartRow_2}";
          Cell cell = sheetData2.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference == cellRef)!;
          string val = GetCellValue(cell);

          student.GetType().GetProperty($"PT{ptCounter}_2")?.SetValue(student, val);
        }

        // Exam_2
        examCellRef = $"{examCol}{rowIndex - _femaleScoresStartRow_1 + _femaleScoresStartRow_2}";
        examCell = sheetData2.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == examCellRef)!;
        student.EX_2 = GetCellValue(examCell);

        femaleChunk.Add(student);        
      }


      if (femaleChunk.Count >= chunkSize) {
        yield return Result.Success((new List<string>(), new List<string>(), exam_1, new List<string>(), new List<string>(), exam_2, new List<StudentWithScores>(), femaleChunk));
        femaleChunk = new List<StudentWithScores>(chunkSize);
      }
    }

    if (femaleChunk.Count > 0)
      yield return Result.Success((new List<string>(), new List<string>(), exam_1, new List<string>(), new List<string>(), exam_2, new List<StudentWithScores>(), femaleChunk));
  }

  public Result SetStudentScore(uint row, int number, string val, ScoreType scoreType = ScoreType.WrittenWorks, bool quarter = true) {
    Result<string> targetColumn = scoreType switch {
      ScoreType.WrittenWorks => Result.Success(ColNumberToName(ColNameToNumber("F")+ (number - 1))),
      ScoreType.PerformanceTasks => Result.Success(ColNumberToName(ColNameToNumber("S")+ (number - 1))),
      ScoreType.Exam => Result.Success(ColNumberToName(ColNameToNumber("AF"))),
      _ => Result.Failure<string>("Invalid ScoreType")
    };

    if (targetColumn.IsFailure) return Result.Failure(targetColumn.Error);

    using var doc = new ECRSpreadSheetDocument();
    var openResult = doc.Open(FilePath, true); if (openResult.IsFailure) return Result.Failure(openResult.Error);
      
    string requiredSheetNames = FirstSem ? requiredSheetNames1[quarter ? 1 : 2] : requiredSheetNames2[quarter ? 1 : 2];

    var sheet = doc.SheetByName(requiredSheetNames);

    if (sheet.IsFailure) return Result.Failure(sheet.Error);

    var wsPartResult = doc.WsPartById(sheet.Value.Id!); if (wsPartResult.IsFailure) return Result.Failure(wsPartResult.Error);
    
    if (wsPartResult.Value is not WorksheetPart wsPart) return Result.Failure("WorksheetPart failure");

    string cellRef = targetColumn.Value + row;
    if (doc.SheetData(wsPart) is not SheetData sheetData)
      return Result.Failure("SheetData is null");

        InsertCellValue(sheetData, cellRef, val);
    
    var workbookResult = doc.Workbook;
    if (workbookResult.IsFailure)
      return Result.Failure(workbookResult.Error);

    var calcProps = workbookResult.Value.CalculationProperties;
    if (calcProps == null) {
      calcProps = new CalculationProperties();
      workbookResult.Value.Append(calcProps);
    }
    calcProps.FullCalculationOnLoad = true;
    calcProps.ForceFullCalculation = true;
    calcProps.CalculationOnSave = true;
    wsPart.Worksheet.Save();
    workbookResult.Value.Save();

    return Result.Success();
  }

  public Result SetHighestScore(int number, string val, ScoreType scoreType = ScoreType.WrittenWorks, bool quarter = true) {
    Result<string> targetColumn = scoreType switch {
      ScoreType.WrittenWorks => Result.Success(ColNumberToName(ColNameToNumber("F")+ (number - 1))),
      ScoreType.PerformanceTasks => Result.Success(ColNumberToName(ColNameToNumber("S")+ (number - 1))),
      ScoreType.Exam => Result.Success(ColNumberToName(ColNameToNumber("AF"))),
      _ => Result.Failure<string>("Invalid ScoreType")
    };

    if (targetColumn.IsFailure) return Result.Failure(targetColumn.Error);

    using var doc = new ECRSpreadSheetDocument();
    var openResult = doc.Open(FilePath, true); if (openResult.IsFailure) return Result.Failure(openResult.Error);
      
    string requiredSheetNames = FirstSem ? requiredSheetNames1[quarter ? 1 : 2] : requiredSheetNames2[quarter ? 1 : 2];

    var sheet = doc.SheetByName(requiredSheetNames);

    if (sheet.IsFailure) return Result.Failure(sheet.Error);

    var wsPartResult = doc.WsPartById(sheet.Value.Id!); if (wsPartResult.IsFailure) return Result.Failure(wsPartResult.Error);
    
    if (wsPartResult.Value is not WorksheetPart wsPart) return Result.Failure("WorksheetPart failure");

    string cellRef = targetColumn.Value + "11";
    if (doc.SheetData(wsPart) is not SheetData sheetData)
      return Result.Failure("SheetData is null");

    InsertCellValue(sheetData, cellRef, val);
    
    var workbookResult = doc.Workbook;
    if (workbookResult.IsFailure)
      return Result.Failure(workbookResult.Error);

    var calcProps = workbookResult.Value.CalculationProperties;
    if (calcProps == null) {
      calcProps = new CalculationProperties();
      workbookResult.Value.Append(calcProps);
    }
    calcProps.FullCalculationOnLoad = true;
    calcProps.ForceFullCalculation = true;
    calcProps.CalculationOnSave = true;
    wsPart.Worksheet.Save();
    workbookResult.Value.Save();

    return Result.Success();
  }

  private static int GetTransmutedGrade(double initialGrade, List<GradeRange> table) {
    var match = table.FirstOrDefault(r => initialGrade >= r.Min && initialGrade <= r.Max);
    return match?.Transmuted ?? 0; // 0 if not found
  }

  public int GetComputedGrade(
    List<string> WrittenWorks, 
    List<string> PerformanceTasks, 
    string Exam,
    List<string> writtenWorksScores, 
    List<string> performanceTaskScores, 
    string examScore
  ) {

    int trackIndex = Array.IndexOf(tracks, Track);

    double writtenWorksTotal = WrittenWorks
        .Select(s => double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var val) ? val : 0)
        .Sum();

    double wwPercentageScore = writtenWorksTotal > 0
        ? writtenWorksScores
            .Select(s => double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var val) ? val : 0)
            .Sum() / writtenWorksTotal * 100.0
        : 0;

    double wwWeightedScore = wwPercentageScore * weightedScores[trackIndex, 0];

    double performanceTasksTotal = PerformanceTasks
        .Select(s => double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var val) ? val : 0)
        .Sum();
    double ptPercentageScore = performanceTasksTotal > 0
        ? performanceTaskScores
            .Select(s => double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var val) ? val : 0)
            .Sum() / performanceTasksTotal * 100.0
        : 0;
    double ptWeightedScore = ptPercentageScore * weightedScores[trackIndex, 1];

    double examTotal = double.TryParse(Exam, NumberStyles.Any, CultureInfo.InvariantCulture, out var et) ? et : 0;
    double examPercentageScore = examTotal > 0
        ? (double.TryParse(examScore, NumberStyles.Any, CultureInfo.InvariantCulture, out var es) ? es : 0) / examTotal * 100.0
        : 0;

    double examWeightedScore = examPercentageScore * weightedScores[trackIndex, 2];

    double initialGrade = wwWeightedScore + ptWeightedScore + examWeightedScore;

    return GetTransmutedGrade(initialGrade, gradeTable);
  }


  private string? ReadCellValue(WorksheetPart wsPart, string cellRef) {
    var cell = wsPart.Worksheet.Descendants<Cell>()
                .FirstOrDefault(c => c.CellReference?.Value == cellRef);
    return cell == null ? null : GetCellValue(cell);
  }

  private static int ColNameToNumber(string colName) {
    int sum = 0;
    foreach (char c in colName.ToUpper()) {
      sum *= 26;
      sum += (c - 'A' + 1);
    }
    return sum;
  }

  private static string ColNumberToName(int colNumber) {
    string colName = "";
    while (colNumber > 0) {
      int rem = (colNumber - 1) % 26;
      colName = (char)('A' + rem) + colName;
      colNumber = (colNumber - 1) / 26;
    }
    return colName;
  }

  private string GetCellValue(Cell cell) {
    if (cell == null || cell.CellValue == null)
      return "";

    if (cell.CellFormula != null)
      return cell.CellValue.InnerText;

    string value = cell.CellValue.InnerText;

    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
      int index = int.Parse(value);
      return _sharedStrings[index];
    }

    return value;
  }

  private static void InsertCellValue(SheetData sheetData, string cellReference, string value, bool isNumeric = true) {
    string rowNumber = new string(cellReference.Where(char.IsDigit).ToArray());
    uint rowIndex = uint.Parse(rowNumber);

    Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
    if (row == null) {
      row = new Row() { RowIndex = rowIndex };
      sheetData.Append(row);
    }

    Cell cell = row.Elements<Cell>()
                   .FirstOrDefault(c => c.CellReference?.Value == cellReference);

    if (cell == null) {
      // Insert in correct order
      Cell refCell = null;
      foreach (Cell existingCell in row.Elements<Cell>()) {
        if (string.Compare(existingCell.CellReference.Value, cellReference, true) > 0) {
          refCell = existingCell;
          break;
        }
      }

      cell = new Cell() { CellReference = cellReference };
      row.InsertBefore(cell, refCell);
    }

    cell.CellValue = new CellValue(value);
    cell.DataType = new EnumValue<CellValues>(isNumeric ? CellValues.Number : CellValues.String);
  }

}
