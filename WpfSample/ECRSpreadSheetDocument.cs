using CSharpFunctionalExtensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WpfSample;

public class ECRSpreadSheetDocument : IDisposable {

    private SpreadsheetDocument _sd;

    public Result Open(string path, bool isEditable = false) { 
        try {
            _sd = SpreadsheetDocument.Open(path, isEditable);            
            return Result.Success();
        } catch (Exception ex) {
            return Result.Failure($"Could not open file: {ex.Message}");
        }
    }

    public Result<WorkbookPart> WorkbookPart => _sd.WorkbookPart == null
        ? Result.Failure<WorkbookPart>("ECR WorkbookPart is null")
        : Result.Success(_sd.WorkbookPart!);

    public Result<Workbook> Workbook => WorkbookPart.IsFailure ?
        Result.Failure<Workbook>(WorkbookPart.Error) :
        Result.Success(WorkbookPart.Value!.Workbook);

    public Result<Sheets> Sheets => Workbook.IsFailure ?
        Result.Failure<Sheets>(Workbook.Error) :
        Workbook.Value.Sheets == null
            ? Result.Failure<Sheets>("ECR Sheets is null")
            : Result.Success(Workbook.Value.Sheets!);

    public Result<Sheet> SheetByName(string name) => Sheets.IsSuccess ?
        Result.Success(
            Sheets.Value!
                .OfType<Sheet>()
                .First(s => s.Name == name)
        ) : 
        Result.Failure<Sheet>(Sheets.Error);

    public Result<WorksheetPart> WsPartById(StringValue sheetId) => WorkbookPart.IsFailure ? 
        Result.Failure<WorksheetPart>(WorkbookPart.Error) :
        Result.Success((WorksheetPart)WorkbookPart.Value.GetPartById(sheetId!));

    public SheetData SheetData(WorksheetPart wsPart) => wsPart.Worksheet.GetFirstChild<SheetData>()!;

    public void Dispose() {
        _sd?.Dispose();
    }

}